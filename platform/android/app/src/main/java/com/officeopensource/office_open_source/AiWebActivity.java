package com.officeopensource.office_open_source;

import android.app.Activity;
import android.content.ClipData;
import android.content.ClipboardManager;
import android.content.Context;
import android.content.Intent;
import android.graphics.Color;
import android.os.Build;
import android.os.Bundle;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.webkit.CookieManager;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.LinearLayout;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

/** v0.0.7: A.I web screen — opens the official website of ChatGPT / Gemini / Copilot / Claude / DeepSeek / Grok / AI Hay inside the app.
 *  Works without login (where the site allows it) or signed in (cookies persist). "Chèn câu trả lời" hands the copied answer back to the document. */
public class AiWebActivity extends Activity {
    WebView web; ProgressBar bar;

    @Override protected void onCreate(Bundle b) {
        super.onCreate(b);
        String url = getIntent().getStringExtra("url"), title = getIntent().getStringExtra("title");
        float dp = getResources().getDisplayMetrics().density;
        LinearLayout root = new LinearLayout(this); root.setOrientation(LinearLayout.VERTICAL);
        LinearLayout top = new LinearLayout(this); top.setOrientation(LinearLayout.HORIZONTAL); top.setGravity(Gravity.CENTER_VERTICAL);
        top.setBackgroundColor(Color.parseColor("#F3F3F3")); top.setPadding((int) (6 * dp), (int) (4 * dp), (int) (6 * dp), (int) (4 * dp));
        Button close = new Button(this); close.setText("← Đóng"); close.setAllCaps(false);
        TextView t = new TextView(this); t.setText(title == null ? "A.I" : title); t.setTextSize(15); t.setSingleLine(true); t.setPadding((int) (8 * dp), 0, (int) (8 * dp), 0);
        Button reload = new Button(this); reload.setText("⟳"); reload.setAllCaps(false);
        Button insert = new Button(this); insert.setText("Chèn câu trả lời"); insert.setAllCaps(false);
        top.addView(close); top.addView(t, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.WRAP_CONTENT, 1f)); top.addView(reload); top.addView(insert);
        bar = new ProgressBar(this, null, android.R.attr.progressBarStyleHorizontal); bar.setMax(100);
        web = new WebView(this);
        root.addView(top); root.addView(bar, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, (int) (3 * dp)));
        root.addView(web, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 0, 1f));
        setContentView(root);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true); s.setDomStorageEnabled(true); s.setDatabaseEnabled(true); s.setLoadWithOverviewMode(true); s.setUseWideViewPort(true);
        s.setSupportMultipleWindows(false); s.setJavaScriptCanOpenWindowsAutomatically(true); s.setMediaPlaybackRequiresUserGesture(true);
        s.setUserAgentString(s.getUserAgentString().replace("; wv", "").replace("Version/4.0 ", "")); // present as Chrome so sign-in pages work
        CookieManager cm = CookieManager.getInstance(); cm.setAcceptCookie(true); if (Build.VERSION.SDK_INT >= 21) cm.setAcceptThirdPartyCookies(web, true);
        web.setWebViewClient(new WebViewClient() {
            @Override public boolean shouldOverrideUrlLoading(WebView v, String u) { return !(u.startsWith("http://") || u.startsWith("https://")); }
        });
        web.setWebChromeClient(new WebChromeClient() { @Override public void onProgressChanged(WebView v, int p) { bar.setProgress(p); bar.setVisibility(p >= 100 ? View.INVISIBLE : View.VISIBLE); } });
        close.setOnClickListener(v -> finish());
        reload.setOnClickListener(v -> web.reload());
        insert.setOnClickListener(v -> {
            String txt = clip();
            if (txt == null || txt.trim().isEmpty()) { Toast.makeText(this, "Hãy bấm nút \"Sao chép\" dưới câu trả lời trên trang trước", Toast.LENGTH_LONG).show(); return; }
            Intent r = new Intent(); r.putExtra("answer", txt); setResult(RESULT_OK, r); finish();
        });
        if (url != null) web.loadUrl(url);
    }
    String clip() { try { ClipboardManager cm = (ClipboardManager) getSystemService(Context.CLIPBOARD_SERVICE); ClipData c = cm.getPrimaryClip(); if (c != null && c.getItemCount() > 0) { CharSequence s = c.getItemAt(0).coerceToText(this); return s == null ? null : s.toString(); } } catch (Exception ignored) {} return null; }
    @Override public void onBackPressed() { if (web.canGoBack()) web.goBack(); else super.onBackPressed(); }
    @Override protected void onPause() { super.onPause(); CookieManager.getInstance().flush(); }
    @Override protected void onDestroy() { if (web != null) web.destroy(); super.onDestroy(); }
}
