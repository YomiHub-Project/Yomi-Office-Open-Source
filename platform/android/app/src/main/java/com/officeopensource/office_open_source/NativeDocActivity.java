package com.officeopensource.office_open_source;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.Rect;
import android.graphics.RectF;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.HandlerThread;
import android.os.Looper;
import android.provider.OpenableColumns;
import android.text.InputType;
import android.util.Log;
import android.util.LruCache;
import android.view.GestureDetector;
import android.view.Gravity;
import android.view.KeyEvent;
import android.view.MotionEvent;
import android.view.ScaleGestureDetector;
import android.view.View;
import android.view.ViewGroup;
import android.view.inputmethod.BaseInputConnection;
import android.view.inputmethod.EditorInfo;
import android.view.inputmethod.InputConnection;
import android.view.inputmethod.InputMethodManager;
import android.widget.Button;
import android.widget.EditText;
import android.widget.HorizontalScrollView;
import android.widget.LinearLayout;
import android.widget.OverScroller;
import android.widget.TextView;
import android.widget.Toast;

import org.libreoffice.kit.Document;
import org.libreoffice.kit.LibreOfficeKit;
import org.libreoffice.kit.Office;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Office Open Source v0.1.6 - native document editor (no WebView).
 * Word / Excel / PowerPoint files are loaded by the embedded LibreOffice core (LibreOfficeKit, MPL-2.0)
 * and rendered as 256px tiles into a native View. Editing uses LOK key/mouse events and .uno: commands.
 * All LibreOfficeKit calls run on one dedicated thread ("lok"), as LibreOffice requires.
 */
public class NativeDocActivity extends Activity {
    static final String TAG = "OOS-LOK";
    static final int TILE = 256;
    static HandlerThread lokThread;
    static Handler lok;
    static Office office;
    static boolean lokReady;

    final Handler ui = new Handler(Looper.getMainLooper());
    Document doc;
    int docType = Document.DOCTYPE_TEXT;
    Uri srcUri; File work; String displayName;
    TileView view;
    LinearLayout partBar; HorizontalScrollView partScroll;
    EditText formula; TextView cellName, status;
    boolean dirty;

    public static boolean available(Context c) {
        try { String[] a = c.getAssets().list("program"); return a != null && a.length > 0; } catch (Exception e) { return false; }
    }

    public static void open(Activity from, Uri uri, String name) {
        Intent i = new Intent(from, NativeDocActivity.class);
        i.setData(uri); i.putExtra("name", name);
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
        if (android.os.Build.VERSION.SDK_INT >= 21) i.addFlags(Intent.FLAG_ACTIVITY_NEW_DOCUMENT | Intent.FLAG_ACTIVITY_MULTIPLE_TASK);
        from.startActivity(i);
    }

    static synchronized void ensureThread() {
        if (lokThread == null) { lokThread = new HandlerThread("lok"); lokThread.start(); lok = new Handler(lokThread.getLooper()); }
    }

    @Override protected void onCreate(Bundle b) {
        super.onCreate(b);
        // viewer mode: never pop up the soft keyboard (the Android core is read-only; editing uses "Chỉnh sửa")
        getWindow().setSoftInputMode(android.view.WindowManager.LayoutParams.SOFT_INPUT_STATE_ALWAYS_HIDDEN);
        srcUri = getIntent().getData();
        displayName = getIntent().getStringExtra("name");
        if (displayName == null) displayName = queryName(srcUri);
        setTitle(displayName);
        if (android.os.Build.VERSION.SDK_INT >= 21) try { setTaskDescription(new android.app.ActivityManager.TaskDescription(displayName)); } catch (Exception ignored) {}
        buildUi();
        status.setText("Đang mở " + displayName + " bằng engine Native (LibreOffice)…");
        ensureThread();
        final Activity act = this;
        lok.post(() -> {
            try {
                if (!lokReady) {
                    unpackAssets(act);   // offapi.rdb, fontconfig, ... must exist as real files (read by non-LO code)
                    LibreOfficeKit.putenv("SAL_LOK_OPTIONS=compact_fonts");
                    LibreOfficeKit.init(act);
                    office = new Office(LibreOfficeKit.getLibreOfficeKitHandle());
                    office.setOptionalFeatures(Document.LOK_FEATURE_DOCUMENT_PASSWORD);
                    lokReady = true;
                }
                work = copyToWork(srcUri, displayName);
                doc = office.documentLoad(work.getAbsolutePath());
                if (doc == null) { final String err = office.getError(); ui.post(() -> fail("Không mở được tệp: " + err)); return; }
                doc.initializeForRendering();
                docType = doc.getDocumentType();

                doc.setMessageCallback((type, payload) -> ui.post(() -> onLok(type, payload)));
                final long w = doc.getDocumentWidth(), h = doc.getDocumentHeight();
                ui.post(() -> { view.setDocSize(w, h); onLoaded(); });
            } catch (Throwable t) {
                Log.e(TAG, "load", t);
                ui.post(() -> fail("Engine Native lỗi: " + t));
            }
        });
    }

    // same as LibreOfficeMainActivity.updatePreferences(): copy assets/unpack/** into the app data dir once per app version
    static void unpackAssets(Activity a) {
        android.content.SharedPreferences sp = a.getSharedPreferences("oos_lok", MODE_PRIVATE);
        long vc = 0; try { vc = a.getPackageManager().getPackageInfo(a.getPackageName(), 0).lastUpdateTime; } catch (Exception ignored) {}
        if (sp.getLong("unpacked2", -1) == vc) return;
        if (copyAssets(a.getAssets(), "unpack", a.getApplicationInfo().dataDir)) sp.edit().putLong("unpacked2", vc).apply();
    }
    static boolean copyAssets(android.content.res.AssetManager am, String from, String to) {
        try {
            String[] files = am.list(from); boolean ok = true;
            for (String f : files) {
                String[] sub = am.list(from + "/" + f);
                if (sub == null || sub.length == 0) {
                    new File(to).mkdirs();
                    try (InputStream in = am.open(from + "/" + f); OutputStream out = new FileOutputStream(to + "/" + f)) {
                        java.io.ByteArrayOutputStream bo = new java.io.ByteArrayOutputStream(); byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) bo.write(buf, 0, n);
                        byte[] data = bo.toByteArray();
                        // fontconfig files carry the LibreOffice Viewer package path: point them at our data dir
                        if (f.endsWith(".conf") || f.endsWith("rc")) data = new String(data, "UTF-8").replace("/data/data/org.documentfoundation.libreoffice", new File(to).getAbsolutePath().replaceAll("/(etc|program|share|user).*$", "").replaceAll("/+$", "")).getBytes("UTF-8");
                        out.write(data);
                    }
                    catch (Exception e) { ok = false; Log.e(TAG, "unpack " + f, e); }
                } else ok &= copyAssets(am, from + "/" + f, to + "/" + f);
            }
            return ok;
        } catch (Exception e) { Log.e(TAG, "unpack", e); return false; }
    }

    void fail(String msg) {
        Toast.makeText(this, msg, Toast.LENGTH_LONG).show();
        status.setText(msg);
    }

    String queryName(Uri u) {
        if (u == null) return "Tài liệu";
        try (Cursor c = getContentResolver().query(u, null, null, null, null)) {
            if (c != null && c.moveToFirst()) { int i = c.getColumnIndex(OpenableColumns.DISPLAY_NAME); if (i >= 0) return c.getString(i); }
        } catch (Exception ignored) {}
        String p = u.getLastPathSegment(); return p != null ? p : "Tài liệu";
    }

    File copyToWork(Uri u, String name) throws Exception {
        // always edit a private copy: LibreOffice opens documents read-only when it cannot create its lock file next to them
        // (shared storage, SD cards, content providers); save() writes the result back to the original location
        File dir = new File(getCacheDir(), "lok/" + Long.toHexString(System.nanoTime())); dir.mkdirs();
        File f = new File(dir, name.replaceAll("[\\\\/:*?\"<>|]", "_"));
        try (InputStream in = "file".equals(u.getScheme()) ? new FileInputStream(u.getPath()) : getContentResolver().openInputStream(u); OutputStream out = new FileOutputStream(f)) {
            byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
        }
        return f;
    }

    // ---------------------------------------------------------------- UI
    float dp(float v) { return v * getResources().getDisplayMetrics().density; }

    Button tb(LinearLayout row, String text, View.OnClickListener l) {
        Button b = new Button(this);
        b.setText(text); b.setAllCaps(false); b.setTextSize(13); b.setMinWidth(0); b.setMinimumWidth(0);
        b.setPadding((int) dp(10), 0, (int) dp(10), 0);
        b.setBackgroundColor(Color.TRANSPARENT); b.setTextColor(0xFF222222);
        b.setOnClickListener(v -> { l.onClick(v); view.requestFocus(); });
        row.addView(b, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, (int) dp(40)));
        return b;
    }
    void uno(String cmd) { uno(cmd, null); }
    void uno(String cmd, String args) { if (doc == null) return; lok.post(() -> doc.postUnoCommand(cmd, args, false)); dirty = true; }

    LinearLayout ribbon, header, tabRow; TextView zoomLabel; int tabSel = 0;
    int accent() { return docType == Document.DOCTYPE_SPREADSHEET ? 0xFF107C41 : docType == Document.DOCTYPE_PRESENTATION ? 0xFFC43E1C : 0xFF185ABD; }
    android.graphics.drawable.GradientDrawable round(int color, float r, int stroke) { android.graphics.drawable.GradientDrawable d = new android.graphics.drawable.GradientDrawable(); d.setColor(color); d.setCornerRadius(dp(r)); if (stroke != 0) d.setStroke((int) dp(1), stroke); return d; }

    void buildUi() {
        LinearLayout root = new LinearLayout(this); root.setOrientation(LinearLayout.VERTICAL); root.setBackgroundColor(0xFFF3F3F3);
        root.setFitsSystemWindows(true);
        header = new LinearLayout(this); header.setOrientation(LinearLayout.VERTICAL); root.addView(header);
        HorizontalScrollView hs = new HorizontalScrollView(this); hs.setHorizontalScrollBarEnabled(false);
        ribbon = new LinearLayout(this); ribbon.setOrientation(LinearLayout.HORIZONTAL); ribbon.setGravity(Gravity.CENTER_VERTICAL);
        ribbon.setPadding((int) dp(6), (int) dp(4), (int) dp(6), (int) dp(4));
        hs.addView(ribbon);
        LinearLayout card = new LinearLayout(this); card.setBackground(round(0xFFFBFBFB, 8, 0xFFE0E0E0));
        card.addView(hs); LinearLayout.LayoutParams cl = new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT); cl.setMargins((int) dp(4), 0, (int) dp(4), (int) dp(4));
        root.addView(card, cl);
        status = new TextView(this); status.setTextSize(12); status.setTextColor(0xFF616161); status.setSingleLine(true);
        view = new TileView(this);
        root.addView(view, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 0, 1f));
        partScroll = new HorizontalScrollView(this); partBar = new LinearLayout(this); partBar.setOrientation(LinearLayout.HORIZONTAL); partBar.setPadding((int) dp(4), (int) dp(2), (int) dp(4), (int) dp(2));
        partScroll.addView(partBar); partScroll.setVisibility(View.GONE); partScroll.setBackgroundColor(0xFFF3F3F3);
        root.addView(partScroll);
        LinearLayout sb = new LinearLayout(this); sb.setOrientation(LinearLayout.HORIZONTAL); sb.setGravity(Gravity.CENTER_VERTICAL); sb.setPadding((int) dp(10), (int) dp(3), (int) dp(10), (int) dp(3));
        zoomLabel = new TextView(this); zoomLabel.setTextSize(12); zoomLabel.setTextColor(0xFF616161);
        sb.addView(status, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1f)); sb.addView(zoomLabel);
        root.addView(sb);
        setContentView(root);
    }

    /** Office-style header: title bar (app badge, file name, search / share / edit) + tab row (Trang đầu / Xem). */
    void buildHeader() {
        header.removeAllViews();
        LinearLayout title = new LinearLayout(this); title.setOrientation(LinearLayout.HORIZONTAL); title.setGravity(Gravity.CENTER_VERTICAL);
        title.setPadding((int) dp(10), (int) dp(6), (int) dp(6), (int) dp(2));
        TextView badge = new TextView(this); badge.setText(docType == Document.DOCTYPE_SPREADSHEET ? "X" : docType == Document.DOCTYPE_PRESENTATION ? "P" : "W");
        badge.setTextColor(Color.WHITE); badge.setTypeface(null, Typeface.BOLD); badge.setGravity(Gravity.CENTER); badge.setTextSize(13); badge.setBackground(round(accent(), 4, 0));
        title.addView(badge, new LinearLayout.LayoutParams((int) dp(24), (int) dp(24)));
        LinearLayout names = new LinearLayout(this); names.setOrientation(LinearLayout.VERTICAL); names.setPadding((int) dp(10), 0, (int) dp(6), 0);
        TextView nm = new TextView(this); nm.setText(displayName); nm.setTextSize(15); nm.setTextColor(0xFF242424); nm.setSingleLine(true); nm.setEllipsize(android.text.TextUtils.TruncateAt.MIDDLE);
        TextView sub = new TextView(this); sub.setText("Chế độ xem • Engine Native"); sub.setTextSize(11); sub.setTextColor(0xFF616161);
        names.addView(nm); names.addView(sub);
        title.addView(names, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1f));
        title.addView(iconBtn("🔍", "Tìm", v -> find()));
        title.addView(iconBtn("📤", "Chia sẻ", v -> exportPdf()));
        Button edit = new Button(this); edit.setText("✏  Chỉnh sửa"); edit.setAllCaps(false); edit.setTextColor(Color.WHITE); edit.setTextSize(13); edit.setTypeface(null, Typeface.BOLD);
        edit.setBackground(round(accent(), 5, 0)); edit.setPadding((int) dp(12), 0, (int) dp(12), 0); edit.setMinWidth(0); edit.setMinimumWidth(0); edit.setMinHeight(0); edit.setMinimumHeight(0);
        edit.setOnClickListener(v -> openInWeb());
        LinearLayout.LayoutParams el = new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, (int) dp(34)); el.setMargins((int) dp(6), 0, 0, 0);
        title.addView(edit, el);
        header.addView(title);
        tabRow = new LinearLayout(this); tabRow.setOrientation(LinearLayout.HORIZONTAL); tabRow.setPadding((int) dp(6), 0, (int) dp(6), 0);
        String[] tabs = { "Trang đầu", "Xem" };
        for (int i = 0; i < tabs.length; i++) {
            final int idx = i; TextView t = new TextView(this); t.setText(tabs[i]); t.setTextSize(14); t.setGravity(Gravity.CENTER); t.setPadding((int) dp(12), (int) dp(8), (int) dp(12), (int) dp(8));
            boolean sel = i == tabSel; t.setTextColor(sel ? accent() : 0xFF242424); if (sel) t.setTypeface(null, Typeface.BOLD);
            if (sel) { android.graphics.drawable.GradientDrawable u = new android.graphics.drawable.GradientDrawable(); u.setColor(Color.TRANSPARENT); android.graphics.drawable.LayerDrawable ld = new android.graphics.drawable.LayerDrawable(new android.graphics.drawable.Drawable[] { new android.graphics.drawable.ColorDrawable(accent()) }); ld.setLayerInset(0, (int) dp(14), (int) dp(33), (int) dp(14), 0); t.setBackground(ld); }
            t.setOnClickListener(v -> { tabSel = idx; buildHeader(); buildRibbon(); });
            tabRow.addView(t, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, (int) dp(36)));
        }
        header.addView(tabRow);
    }
    View iconBtn(String glyph, String desc, View.OnClickListener l) {
        TextView t = new TextView(this); t.setText(glyph); t.setTextSize(18); t.setGravity(Gravity.CENTER); t.setContentDescription(desc);
        t.setBackground(round(Color.TRANSPARENT, 18, 0)); t.setOnClickListener(l);
        return withSize(t, 40, 40);
    }
    View withSize(View v, int w, int h) { v.setLayoutParams(new LinearLayout.LayoutParams((int) dp(w), (int) dp(h))); return v; }

    /** Large ribbon command: icon above label, like the desktop Ribbon. */
    View big(String glyph, String label, View.OnClickListener l) {
        LinearLayout b = new LinearLayout(this); b.setOrientation(LinearLayout.VERTICAL); b.setGravity(Gravity.CENTER_HORIZONTAL);
        b.setPadding((int) dp(8), (int) dp(4), (int) dp(8), (int) dp(4)); b.setClickable(true);
        android.graphics.drawable.StateListDrawable sd = new android.graphics.drawable.StateListDrawable();
        sd.addState(new int[] { android.R.attr.state_pressed }, round(0xFFE0E0E0, 6, 0)); sd.addState(new int[0], round(Color.TRANSPARENT, 6, 0)); b.setBackground(sd);
        TextView ic = new TextView(this); ic.setText(glyph); ic.setTextSize(22); ic.setGravity(Gravity.CENTER);
        TextView tx = new TextView(this); tx.setText(label); tx.setTextSize(12); tx.setTextColor(0xFF242424); tx.setGravity(Gravity.CENTER);
        b.addView(ic); b.addView(tx);
        b.setOnClickListener(v -> { l.onClick(v); view.requestFocus(); });
        ribbon.addView(b, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, ViewGroup.LayoutParams.WRAP_CONTENT));
        return b;
    }
    void sep() { View v = new View(this); v.setBackgroundColor(0xFFE0E0E0); LinearLayout.LayoutParams lp = new LinearLayout.LayoutParams((int) dp(1), (int) dp(40)); lp.setMargins((int) dp(4), 0, (int) dp(4), 0); ribbon.addView(v, lp); }

    // The Android LibreOfficeKit core shipped here (LibreOffice Viewer build) renders documents faithfully but keeps the
    // view read-only: key events, .uno:InsertText and paste are ignored. Android therefore uses the native engine as a
    // high-fidelity viewer and hands editing to the built-in Office Open Source editor ("Chỉnh sửa").
    void buildRibbon() {
        ribbon.removeAllViews();
        if (tabSel == 0) {
            big("✏️", "Chỉnh sửa", v -> openInWeb());
            sep();
            big("🔍", "Tìm", v -> find());
            big("📄", "Xuất PDF", v -> exportPdf());
            big("📤", "Chia sẻ", v -> exportPdf());
            if (docType == Document.DOCTYPE_PRESENTATION) { sep(); big("▶️", "Trình chiếu", v -> slideShow()); }
            sep();
            big("ℹ️", "Thông tin", v -> info());
        } else {
            big("➕", "Phóng to", v -> { view.zoomBy(1.25f); updateZoom(); });
            big("➖", "Thu nhỏ", v -> { view.zoomBy(0.8f); updateZoom(); });
            big("↔️", "Vừa trang", v -> { view.fitWidth(); updateZoom(); });
            if (docType == Document.DOCTYPE_PRESENTATION) { sep(); big("▶️", "Trình chiếu", v -> slideShow()); }
        }
    }
    void updateZoom() { if (zoomLabel != null) zoomLabel.setText(Math.round(view.zoomPercent()) + "%"); }
    void info() {
        new AlertDialog.Builder(this).setTitle(displayName)
            .setMessage("Hiển thị bằng engine Native (LibreOffice) – chuẩn bố cục Microsoft Office.\nBấm \"Chỉnh sửa\" để sửa nội dung bằng trình soạn thảo Office Open Source.")
            .setPositiveButton("OK", null).show();
    }

    void buildFormulaBar() {
        LinearLayout bar = new LinearLayout(this); bar.setOrientation(LinearLayout.HORIZONTAL); bar.setBackgroundColor(Color.WHITE);
        cellName = new TextView(this); cellName.setMinWidth((int) dp(64)); cellName.setPadding((int) dp(8), 0, (int) dp(8), 0); cellName.setGravity(Gravity.CENTER_VERTICAL);
        TextView fx = new TextView(this); fx.setText("fx"); fx.setTypeface(Typeface.SERIF, Typeface.ITALIC); fx.setPadding((int) dp(6), 0, (int) dp(6), 0); fx.setGravity(Gravity.CENTER_VERTICAL);
        formula = new EditText(this); formula.setSingleLine(true); formula.setTextSize(14); formula.setImeOptions(EditorInfo.IME_ACTION_DONE);
        formula.setOnEditorActionListener((v, id, ev) -> {
            final String t = formula.getText().toString();
            lok.post(() -> { key(1286, 0); for (int i = 0; i < t.length(); i++) key(0, t.charAt(i)); key(1280, 13); });
            dirty = true; view.requestFocus(); return true;
        });
        bar.addView(cellName, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, ViewGroup.LayoutParams.MATCH_PARENT));
        bar.addView(fx, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, ViewGroup.LayoutParams.MATCH_PARENT));
        bar.addView(formula, new LinearLayout.LayoutParams(0, (int) dp(44), 1f));
        ((LinearLayout) view.getParent()).addView(bar, 1);
    }

    void key(int awt, int ch) { Log.d(TAG, "postKeyEvent awt=" + awt + " ch=" + ch); doc.postKeyEvent(Document.KEY_EVENT_PRESS, ch, awt); doc.postKeyEvent(Document.KEY_EVENT_RELEASE, ch, awt); }

    void onLoaded() {
        buildHeader(); buildRibbon();
        if (docType == Document.DOCTYPE_SPREADSHEET) buildFormulaBar();
        refreshParts();
        String kind = docType == Document.DOCTYPE_TEXT ? "Word" : docType == Document.DOCTYPE_SPREADSHEET ? "Excel" : docType == Document.DOCTYPE_PRESENTATION ? "PowerPoint" : "Draw";
        status.setText(kind + " • Chế độ xem • Engine Native (LibreOffice)");
        view.fitWidth(); updateZoom();
    }

    void refreshParts() {
        if (doc == null || (docType != Document.DOCTYPE_SPREADSHEET && docType != Document.DOCTYPE_PRESENTATION)) return;
        lok.post(() -> {
            final int n = doc.getParts(), cur = doc.getPart();
            final List<String> names = new ArrayList<>();
            for (int i = 0; i < n; i++) { String s = doc.getPartName(i); names.add(s == null || s.isEmpty() ? (docType == Document.DOCTYPE_SPREADSHEET ? "Sheet" : "Slide ") + (i + 1) : s); }
            ui.post(() -> {
                partBar.removeAllViews(); partScroll.setVisibility(View.VISIBLE);
                for (int i = 0; i < n; i++) {
                    final int idx = i;
                    Button b = tb(partBar, (docType == Document.DOCTYPE_PRESENTATION ? (i + 1) + ". " : "") + names.get(i), v -> setPart(idx));
                    if (i == cur) { b.setBackground(round(Color.WHITE, 4, 0)); b.setTextColor(accent()); b.setTypeface(null, Typeface.BOLD); }
                }
            });
        });
    }

    void setPart(int i) {
        lok.post(() -> { doc.setPart(i); final long w = doc.getDocumentWidth(), h = doc.getDocumentHeight(); ui.post(() -> { view.setDocSize(w, h); view.dropAll(); view.scrollToTop(); refreshParts(); }); });
    }

    void onLok(int type, String p) {
        if (type == Document.CALLBACK_STATE_CHANGED && p != null && (p.contains("EditDoc") || p.contains("ReadOnly") || p.contains("InsertText"))) Log.d(TAG, "state " + p);
        if (type == Document.CALLBACK_ERROR) Log.w(TAG, "lok error " + p);
        switch (type) {
            case Document.CALLBACK_INVALIDATE_TILES: view.invalidateTw(p); break;
            case Document.CALLBACK_INVALIDATE_VISIBLE_CURSOR: view.caret = parse(p); view.ensureVisible(view.caret); view.invalidate(); break;
            case Document.CALLBACK_TEXT_SELECTION: view.selection.clear(); if (p != null && !p.isEmpty()) for (String s : p.split(";")) { RectF r = parse(s); if (r != null) view.selection.add(r); } view.invalidate(); break;
            case Document.CALLBACK_CURSOR_VISIBLE: view.caretVisible = "true".equals(p); view.invalidate(); break;
            case Document.CALLBACK_CELL_CURSOR: view.cell = "EMPTY".equals(p) ? null : parse(p); if (view.cell != null) view.ensureVisible(view.cell); view.invalidate(); break;
            case Document.CALLBACK_DOCUMENT_SIZE_CHANGED: lok.post(() -> { final long w = doc.getDocumentWidth(), h = doc.getDocumentHeight(); ui.post(() -> { view.setDocSize(w, h); view.dropAll(); }); }); break;
            case Document.CALLBACK_SET_PART: refreshParts(); break;
            case Document.CALLBACK_CELL_FORMULA: if (formula != null && !formula.hasFocus()) formula.setText(p); break;
            case Document.CALLBACK_CELL_ADDRESS: if (cellName != null) cellName.setText(p); break;
            case Document.CALLBACK_HYPERLINK_CLICKED: if (p != null && (p.startsWith("http") || p.startsWith("mailto:"))) try { startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(p))); } catch (Exception ignored) {} break;
            case Document.CALLBACK_STATE_CHANGED: if (p != null && p.startsWith(".uno:ModifiedStatus=")) dirty = p.endsWith("true"); break;
        }
    }

    static RectF parse(String s) {
        if (s == null) return null;
        String[] a = s.split(",");
        if (a.length < 4) return null;
        try { float x = Float.parseFloat(a[0].trim()), y = Float.parseFloat(a[1].trim()), w = Float.parseFloat(a[2].trim()), h = Float.parseFloat(a[3].trim()); return new RectF(x, y, x + w, y + h); }
        catch (Exception e) { return null; }
    }

    // ---------------------------------------------------------------- file ops
    static String ext(String n) { int i = n.lastIndexOf('.'); return i < 0 ? "" : n.substring(i + 1).toLowerCase(); }

    void save() {
        if (doc == null) return;
        final String e = ext(displayName);
        status.setText("Đang lưu…");
        lok.post(() -> {
            try {
                File tmp = new File(work.getParentFile(), "saving." + e);
                doc.saveAs(tmp.getAbsolutePath(), e, "");
                if (!"file".equals(srcUri.getScheme())) {
                    try (InputStream in = new FileInputStream(tmp); OutputStream out = getContentResolver().openOutputStream(srcUri, "wt")) { byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) out.write(buf, 0, n); }
                } else {
                    try (InputStream in = new FileInputStream(tmp); OutputStream out = new FileOutputStream(srcUri.getPath())) { byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) out.write(buf, 0, n); }
                }
                tmp.delete();
                dirty = false;
                ui.post(() -> { status.setText("Đã lưu " + displayName); Toast.makeText(this, "Đã lưu", Toast.LENGTH_SHORT).show(); });
            } catch (Exception ex) { ui.post(() -> fail("Không lưu được: " + ex.getMessage())); }
        });
    }

    void exportPdf() {
        if (doc == null) return;
        lok.post(() -> {
            try {
                File dir = new File(getCacheDir(), "share"); dir.mkdirs();
                File pdf = new File(dir, displayName.replaceAll("\\.[^.]*$", "") + ".pdf");
                doc.saveAs(pdf.getAbsolutePath(), "pdf", "");
                ui.post(() -> {
                    try {
                        Uri u = Uri.parse("content://com.officeopensource.office_open_source.share/" + Uri.encode(pdf.getName()));
                        Intent s = new Intent(Intent.ACTION_SEND).setType("application/pdf").putExtra(Intent.EXTRA_STREAM, u).addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                        startActivity(Intent.createChooser(s, "Chia sẻ / lưu PDF"));
                    } catch (Exception ex) { fail("PDF: " + ex.getMessage()); }
                });
            } catch (Exception ex) { ui.post(() -> fail("Xuất PDF lỗi: " + ex.getMessage())); }
        });
    }

    void openInWeb() {
        Intent i = new Intent(this, MainActivity.class); i.setAction(Intent.ACTION_VIEW); i.setData(srcUri);
        i.putExtra("forceWeb", true); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
        startActivity(i);
    }

    void find() {
        final EditText t = new EditText(this); t.setHint("Từ cần tìm");
        new AlertDialog.Builder(this).setTitle("Tìm trong tài liệu").setView(t)
            .setPositiveButton("Tìm tiếp", (d, w) -> uno(".uno:ExecuteSearch", "{\"SearchItem.SearchString\":{\"type\":\"string\",\"value\":\"" + t.getText().toString().replace("\\", "\\\\").replace("\"", "\\\"") + "\"},\"SearchItem.Backward\":{\"type\":\"boolean\",\"value\":false},\"SearchItem.Command\":{\"type\":\"long\",\"value\":0}}"))
            .setNegativeButton("Đóng", null).show();
    }

    void runMacro() {
        final EditText t = new EditText(this); t.setText("Module1."); t.setInputType(InputType.TYPE_CLASS_TEXT);
        new AlertDialog.Builder(this).setTitle("Chạy macro VBA / Basic").setMessage("Ví dụ: Module1.TenMacro").setView(t)
            .setPositiveButton("Chạy", (d, w) -> {
                String n = t.getText().toString().trim();
                String url = n.startsWith("vnd.sun.star.script:") ? n : (n.split("\\.").length == 2 ? "vnd.sun.star.script:VBAProject." + n : "vnd.sun.star.script:" + n) + "?language=Basic&location=document";
                uno(url);
            }).setNegativeButton("Hủy", null).show();
    }

    void slideShow() {
        if (doc == null) return;
        final int[] cur = { 0 };
        final android.widget.ImageView iv = new android.widget.ImageView(this); iv.setBackgroundColor(Color.BLACK); iv.setScaleType(android.widget.ImageView.ScaleType.FIT_CENTER);
        final android.app.Dialog dlg = new android.app.Dialog(this, android.R.style.Theme_Black_NoTitleBar_Fullscreen); dlg.setContentView(iv);
        final int orig[] = { 0 };
        final Runnable render = () -> lok.post(() -> {
            doc.setPart(cur[0]);
            long tw = doc.getDocumentWidth(), th = doc.getDocumentHeight(); if (tw <= 0) return;
            int w = getResources().getDisplayMetrics().widthPixels, h = (int) (w * th / tw);
            ByteBuffer buf = ByteBuffer.allocateDirect(w * h * 4);
            doc.paintTile(buf, w, h, 0, 0, (int) tw, (int) th);
            final Bitmap bm = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888); buf.rewind(); bm.copyPixelsFromBuffer(buf);
            ui.post(() -> iv.setImageBitmap(bm));
        });
        lok.post(() -> { orig[0] = doc.getPart(); final int n = doc.getParts(); ui.post(() -> {
            iv.setOnClickListener(v -> { if (cur[0] < n - 1) { cur[0]++; render.run(); } else dlg.dismiss(); });
            iv.setOnLongClickListener(v -> { if (cur[0] > 0) { cur[0]--; render.run(); } return true; });
            dlg.setOnDismissListener(d -> setPart(orig[0]));
            dlg.show(); render.run();
        }); });
    }

    @Override public void onBackPressed() {
        if (!dirty) { super.onBackPressed(); return; }
        new AlertDialog.Builder(this).setTitle(displayName).setMessage("Lưu thay đổi trước khi đóng?")
            .setPositiveButton("Lưu", (d, w) -> { save(); ui.postDelayed(this::finish, 1200); })
            .setNegativeButton("Không lưu", (d, w) -> finish())
            .setNeutralButton("Hủy", null).show();
    }

    @Override protected void onDestroy() {
        final Document d = doc; doc = null;
        if (d != null) lok.post(d::destroy);
        view.dropAll();
        super.onDestroy();
    }

    // ---------------------------------------------------------------- tiled view
    class TileView extends View {
        // v0.1.7: mouse pointer on DeX / external mouse - Excel cell cross on worksheets, text I-beam in Word, arrow in PowerPoint
        @android.annotation.TargetApi(24)
        @Override public android.view.PointerIcon onResolvePointerIcon(MotionEvent e, int pointerIndex) {
            int t = docType == Document.DOCTYPE_SPREADSHEET ? android.view.PointerIcon.TYPE_CELL : docType == Document.DOCTYPE_TEXT ? android.view.PointerIcon.TYPE_TEXT : android.view.PointerIcon.TYPE_ARROW;
            return android.view.PointerIcon.getSystemIcon(getContext(), t);
        }
        float scale = 0.05f;                 // device px per twip
        float sx, sy;                        // scroll in px
        long docW = 1, docH = 1;             // twips
        final LruCache<Long, Bitmap> tiles;
        final Set<Long> pending = new HashSet<>();
        final Paint paint = new Paint(Paint.FILTER_BITMAP_FLAG), sel = new Paint(), caretP = new Paint(), cellP = new Paint(), shadow = new Paint();
        RectF caret, cell; boolean caretVisible = true; final List<RectF> selection = new ArrayList<>();
        final OverScroller scroller;
        final GestureDetector gd; final ScaleGestureDetector sg;
        int gen = 0;

        TileView(Context c) {
            super(c);
            setFocusable(true); setFocusableInTouchMode(true);
            setBackgroundColor(0xFFE6E6E6);
            int mem = (int) Math.min(Runtime.getRuntime().maxMemory() / 4, 96L * 1024 * 1024);
            tiles = new LruCache<Long, Bitmap>(mem) { @Override protected int sizeOf(Long k, Bitmap b) { return b.getByteCount(); } };
            sel.setColor(0x552173D9); caretP.setColor(Color.BLACK); cellP.setColor(0xFF107C41); cellP.setStyle(Paint.Style.STROKE); cellP.setStrokeWidth(dp(2));
            shadow.setColor(0x33000000);
            scroller = new OverScroller(c);
            gd = new GestureDetector(c, new GestureDetector.SimpleOnGestureListener() {
                @Override public boolean onDown(MotionEvent e) { scroller.forceFinished(true); return true; }
                @Override public boolean onScroll(MotionEvent a, MotionEvent b, float dx, float dy) { scrollBy_(dx, dy); return true; }
                @Override public boolean onFling(MotionEvent a, MotionEvent b, float vx, float vy) { scroller.fling((int) sx, (int) sy, (int) -vx, (int) -vy, 0, maxX(), 0, maxY()); postInvalidateOnAnimation(); return true; }
                @Override public boolean onSingleTapUp(MotionEvent e) { click(e.getX(), e.getY(), 1, Document.MOUSE_BUTTON_LEFT); return true; }
                @Override public boolean onDoubleTap(MotionEvent e) { click(e.getX(), e.getY(), 2, Document.MOUSE_BUTTON_LEFT); return true; }
                @Override public void onLongPress(MotionEvent e) { final float x = tw(e.getX(), true), y = tw(e.getY(), false); if (doc != null) lok.post(() -> doc.setTextSelection(Document.SET_TEXT_SELECTION_START, (int) x, (int) y)); }
            });
            sg = new ScaleGestureDetector(c, new ScaleGestureDetector.SimpleOnScaleGestureListener() {
                @Override public boolean onScale(ScaleGestureDetector d) { zoomAt(d.getScaleFactor(), d.getFocusX(), d.getFocusY(), false); return true; }
                @Override public void onScaleEnd(ScaleGestureDetector d) { zoomAt(1f, d.getFocusX(), d.getFocusY(), true); }
            });
        }

        float liveScale = 1f; // visual scale during pinch (tiles re-rendered at the end)
        void zoomAt(float f, float fx, float fy, boolean commit) {
            if (!commit) { liveScale *= f; invalidate(); return; }
            float ns = Math.max(0.005f, Math.min(0.5f, scale * liveScale));
            float docX = (fx + sx - ox()) / scale, docY = (fy + sy - oy()) / scale;
            scale = ns; liveScale = 1f; dropAll();
            sx = docX * scale - fx + ox(); sy = docY * scale - fy + oy(); clamp();
            invalidate();
        }
        void zoomBy(float f) { zoomAt(f, getWidth() / 2f, getHeight() / 2f, false); zoomAt(1f, getWidth() / 2f, getHeight() / 2f, true); }
        float zoomPercent() { return scale / (getResources().getDisplayMetrics().density * 96f / 1440f) * 100f; }
        void fitWidth() {
            if (getWidth() == 0) { post(() -> { fitWidth(); updateZoom(); }); return; }
            long w = docType == Document.DOCTYPE_SPREADSHEET ? Math.min(docW, 12000) : Math.min(docW, 12240 * 2);
            scale = (getWidth() - dp(16)) / (float) Math.max(1, w);
            dropAll(); invalidate();
        }
        void setDocSize(long w, long h) { docW = Math.max(1, w); docH = Math.max(1, h); clamp(); invalidate(); }
        void scrollToTop() { sx = 0; sy = 0; invalidate(); }
        float ox() { float dw = docW * scale; return dw < getWidth() ? (getWidth() - dw) / 2f : 0; }
        float oy() { return dp(8); }
        int maxX() { return (int) Math.max(0, docW * scale - getWidth() + dp(8)); }
        int maxY() { return (int) Math.max(0, docH * scale - getHeight() + dp(48)); }
        void clamp() { sx = Math.max(0, Math.min(sx, maxX())); sy = Math.max(0, Math.min(sy, maxY())); }
        void scrollBy_(float dx, float dy) { sx += dx; sy += dy; clamp(); invalidate(); }
        float tw(float px, boolean x) { return x ? (px + sx - ox()) / scale : (px + sy - oy()) / scale; }
        float px(float t, boolean x) { return x ? t * scale - sx + ox() : t * scale - sy + oy(); }
        RectF toPx(RectF r) { return new RectF(px(r.left, true), px(r.top, false), px(r.right, true), px(r.bottom, false)); }

        void ensureVisible(RectF r) {
            if (r == null) return; RectF p = toPx(r);
            if (p.top < 0) sy += p.top - dp(40); else if (p.bottom > getHeight()) sy += p.bottom - getHeight() + dp(40);
            if (p.left < 0) sx += p.left - dp(20); else if (p.right > getWidth()) sx += p.right - getWidth() + dp(20);
            clamp();
        }

        void click(float x, float y, int count, int btn) {
            if (doc == null) return;
            final int tx = (int) tw(x, true), ty = (int) tw(y, false);
            final int tt = (int) (TILE / scale);
            lok.post(() -> {
                doc.setClientZoom(TILE, TILE, tt, tt);
                doc.postMouseEvent(Document.MOUSE_EVENT_BUTTON_DOWN, tx, ty, count, btn, Document.KEYBOARD_MODIFIER_NONE);
                doc.postMouseEvent(Document.MOUSE_EVENT_BUTTON_UP, tx, ty, count, btn, Document.KEYBOARD_MODIFIER_NONE);
            });
            requestFocus();
        }

        void dropAll() { gen++; tiles.evictAll(); synchronized (pending) { pending.clear(); } }
        void invalidateTw(String p) {
            if (p == null || p.startsWith("EMPTY")) { dropAll(); invalidate(); return; }
            RectF r = parse(p); if (r == null) return;
            float t = TILE / scale;
            for (Long k : new ArrayList<>(tiles.snapshot().keySet())) {
                int tx = (int) (k >> 32), ty = (int) (long) (k & 0xffffffffL);
                if (RectF.intersects(r, new RectF(tx * t, ty * t, (tx + 1) * t, (ty + 1) * t))) tiles.remove(k);
            }
            invalidate();
        }

        void request(final int tx, final int ty) {
            final long k = ((long) tx << 32) | (ty & 0xffffffffL);
            synchronized (pending) { if (pending.contains(k)) return; pending.add(k); }
            final int g = gen; final float sc = scale;
            lok.post(() -> {
                if (g != gen || doc == null) { synchronized (pending) { pending.remove(k); } return; }
                int t = (int) (TILE / sc);
                ByteBuffer buf = ByteBuffer.allocateDirect(TILE * TILE * 4);
                doc.paintTile(buf, TILE, TILE, tx * t, ty * t, t, t);
                Bitmap bm = Bitmap.createBitmap(TILE, TILE, Bitmap.Config.ARGB_8888);
                buf.rewind(); bm.copyPixelsFromBuffer(buf);
                ui.post(() -> { synchronized (pending) { pending.remove(k); } if (g == gen) { tiles.put(k, bm); invalidate(); } });
            });
        }

        @Override public void computeScroll() { if (scroller.computeScrollOffset()) { sx = scroller.getCurrX(); sy = scroller.getCurrY(); clamp(); postInvalidateOnAnimation(); } }

        @Override protected void onDraw(Canvas c) {
            if (doc == null) return;
            c.save();
            if (liveScale != 1f) c.scale(liveScale, liveScale, getWidth() / 2f, getHeight() / 2f);
            float ox = ox() - sx, oy = oy() - sy; float dw = docW * scale, dh = docH * scale;
            c.drawRect(ox + dp(2), oy + dp(2), ox + dw + dp(2), oy + dh + dp(2), shadow);
            Paint white = new Paint(); white.setColor(Color.WHITE); c.drawRect(ox, oy, ox + dw, oy + dh, white);
            int x0 = (int) Math.max(0, Math.floor(-ox / TILE)), y0 = (int) Math.max(0, Math.floor(-oy / TILE));
            int x1 = (int) Math.min(Math.ceil(dw / TILE) - 1, Math.floor((getWidth() - ox) / TILE));
            int y1 = (int) Math.min(Math.ceil(dh / TILE) - 1, Math.floor((getHeight() - oy) / TILE));
            for (int ty = y0; ty <= y1; ty++) for (int tx = x0; tx <= x1; tx++) {
                Bitmap b = tiles.get(((long) tx << 32) | (ty & 0xffffffffL));
                if (b != null) c.drawBitmap(b, ox + tx * TILE, oy + ty * TILE, paint); else request(tx, ty);
            }
            for (RectF r : selection) c.drawRect(toPx(r), sel);
            if (cell != null) c.drawRect(toPx(cell), cellP);
            if (caret != null && caretVisible && hasFocus()) { RectF p = toPx(caret); c.drawRect(p.left, p.top, p.left + Math.max(dp(2), p.width()), p.bottom, caretP); }
            c.restore();
        }

        @Override public boolean onTouchEvent(MotionEvent e) {
            sg.onTouchEvent(e);
            if (!sg.isInProgress()) gd.onTouchEvent(e);
            if (e.isFromSource(android.view.InputDevice.SOURCE_MOUSE) && e.getActionMasked() == MotionEvent.ACTION_MOVE && doc != null) {
                final int tx = (int) tw(e.getX(), true), ty = (int) tw(e.getY(), false);
                lok.post(() -> doc.postMouseEvent(Document.MOUSE_EVENT_MOVE, tx, ty, 1, Document.MOUSE_BUTTON_LEFT, 0));
            }
            return true;
        }
        @Override public boolean onGenericMotionEvent(MotionEvent e) {
            if (e.getActionMasked() == MotionEvent.ACTION_SCROLL) {
                float v = e.getAxisValue(MotionEvent.AXIS_VSCROLL);
                if ((e.getMetaState() & KeyEvent.META_CTRL_ON) != 0) zoomBy(v > 0 ? 1.1f : 0.9f); else scrollBy_(0, -v * dp(64));
                return true;
            }
            return super.onGenericMotionEvent(e);
        }

        void showKeyboard() {
            requestFocus();
            InputMethodManager im = (InputMethodManager) getSystemService(Context.INPUT_METHOD_SERVICE);
            if (im != null) im.showSoftInput(this, 0);
        }

        @Override public boolean onCheckIsTextEditor() { return false; }
        @Override public InputConnection onCreateInputConnection(EditorInfo o) {
            o.inputType = InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_FLAG_MULTI_LINE;
            o.imeOptions = EditorInfo.IME_FLAG_NO_EXTRACT_UI | EditorInfo.IME_FLAG_NO_FULLSCREEN;
            return new BaseInputConnection(this, false) {
                @Override public boolean commitText(CharSequence t, int pos) { typeText(t.toString()); return true; }
                @Override public boolean deleteSurroundingText(int before, int after) { for (int i = 0; i < before; i++) post_(1283, 0); for (int i = 0; i < after; i++) post_(1286, 0); return true; }
                @Override public boolean sendKeyEvent(KeyEvent e) { return e.getAction() == KeyEvent.ACTION_DOWN ? onKeyDown(e.getKeyCode(), e) : true; }
                @Override public boolean setComposingText(CharSequence t, int p) { return true; } // committed by commitText (Telex/VNI IMEs commit per word)
                @Override public boolean finishComposingText() { return true; }
            };
        }
        void post_(int awt, int ch) {
            if (doc == null) return;
            final boolean text = docType != Document.DOCTYPE_SPREADSHEET;
            if (text && awt == 1283) lok.post(() -> doc.postUnoCommand(docType == Document.DOCTYPE_TEXT ? ".uno:SwBackspace" : ".uno:Backspace", null, false));
            else if (text && awt == 1286) lok.post(() -> doc.postUnoCommand(".uno:Delete", null, false));
            else lok.post(() -> key(awt, ch));
            dirty = true;
        }
        void typeText(String s) {
            Log.d(TAG, "typeText '" + s + "' doc=" + (doc != null));
            if (doc == null) return;
            // text goes in through .uno:InsertText (reliable on the Android LibreOfficeKit build, where plain key
            // events reach the view but are not inserted); Enter / Excel cell typing still use key events
            lok.post(() -> {
                StringBuilder run = new StringBuilder();
                for (int i = 0; i <= s.length(); ) {
                    int cp = i < s.length() ? s.codePointAt(i) : -1;
                    if (cp == -1 || cp == '\n') {
                        if (run.length() > 0) {
                            if (docType == Document.DOCTYPE_SPREADSHEET) { for (int k = 0; k < run.length(); ) { int c = run.codePointAt(k); k += Character.charCount(c); key(0, c); } }
                            else if (!doc.paste("text/plain;charset=utf-8", run.toString())) doc.postUnoCommand(".uno:InsertText", "{\"Text\":{\"type\":\"string\",\"value\":\"" + run.toString().replace("\\", "\\\\").replace("\"", "\\\"") + "\"}}", false);
                            run.setLength(0);
                        }
                        if (cp == '\n') key(1280, 13);
                        if (cp == -1) break;
                    } else run.appendCodePoint(cp);
                    i += Character.charCount(Math.max(cp, 0));
                }
            });
            dirty = true;
        }

        @Override public boolean onKeyDown(int code, KeyEvent e) {
            Log.d(TAG, "onKeyDown " + code);
            int awt = 0; int mods = 0;
            if (e.isShiftPressed()) mods |= Document.KEYBOARD_MODIFIER_SHIFT; if (e.isCtrlPressed()) mods |= Document.KEYBOARD_MODIFIER_MOD1; if (e.isAltPressed()) mods |= Document.KEYBOARD_MODIFIER_MOD2;
            switch (code) {
                case KeyEvent.KEYCODE_DEL: awt = 1283; break; case KeyEvent.KEYCODE_FORWARD_DEL: awt = 1286; break;
                case KeyEvent.KEYCODE_ENTER: case KeyEvent.KEYCODE_NUMPAD_ENTER: awt = 1280; break; case KeyEvent.KEYCODE_TAB: awt = 1282; break;
                case KeyEvent.KEYCODE_DPAD_DOWN: awt = 1024; break; case KeyEvent.KEYCODE_DPAD_UP: awt = 1025; break;
                case KeyEvent.KEYCODE_DPAD_LEFT: awt = 1026; break; case KeyEvent.KEYCODE_DPAD_RIGHT: awt = 1027; break;
                case KeyEvent.KEYCODE_MOVE_HOME: awt = 1028; break; case KeyEvent.KEYCODE_MOVE_END: awt = 1029; break;
                case KeyEvent.KEYCODE_PAGE_UP: awt = 1030; break; case KeyEvent.KEYCODE_PAGE_DOWN: awt = 1031; break;
                case KeyEvent.KEYCODE_ESCAPE: awt = 1281; break;
            }
            if (e.isCtrlPressed() && code >= KeyEvent.KEYCODE_A && code <= KeyEvent.KEYCODE_Z) {
                if (code == KeyEvent.KEYCODE_S) { save(); return true; }
                awt = 512 + (code - KeyEvent.KEYCODE_A);
            }
            if (awt != 0) { final int k = awt | mods; post_(k, awt == 1280 ? 13 : 0); return true; }
            int uc = e.getUnicodeChar(e.getMetaState());
            if (uc > 0) { typeText(new String(Character.toChars(uc))); return true; }
            return super.onKeyDown(code, e);
        }
    }
}
