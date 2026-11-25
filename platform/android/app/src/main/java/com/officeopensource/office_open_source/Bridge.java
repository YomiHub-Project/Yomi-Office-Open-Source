package com.officeopensource.office_open_source;

import android.app.ActivityManager;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.print.PrintAttributes;
import android.print.PrintManager;
import android.provider.OpenableColumns;
import android.util.Base64;
import android.webkit.JavascriptInterface;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/** JS <-> Android bridge exposed as window.OOSNative (see src/core/host.js). Async methods answer via __oosCb(id, json). */
public class Bridge {
    final MainActivity a;
    String pending;            // JSON {name,b64,uri} of a file opened from another app
    byte[] saveBytes; String saveName, saveId;

    Bridge(MainActivity a) { this.a = a; }

    // v0.0.6: files handed to a brand-new window (DocActivity task) — same process, so a static map is enough
    static final java.util.Map<String, String> STASH = new java.util.HashMap<>();

    boolean takeIntent(Intent i) {
        if (i == null) return false;
        String tok = i.getStringExtra("stash");
        if (tok != null) { synchronized (STASH) { String v = STASH.remove(tok); if (v != null) { pending = v; return true; } } }
        Uri u = null;
        if (Intent.ACTION_VIEW.equals(i.getAction()) || Intent.ACTION_EDIT.equals(i.getAction())) u = i.getData();
        else if (Intent.ACTION_SEND.equals(i.getAction())) u = i.getParcelableExtra(Intent.EXTRA_STREAM);
        if (u == null) return false;
        // v0.1.6: Word/Excel/PowerPoint files open in the native LibreOfficeKit editor (no WebView) unless the user chose the web engine
        boolean forceWeb = i.getBooleanExtra("forceWeb", false);
        if (!forceWeb && nativeEngineWanted() && isOfficeName(nameOf(u))) {
            try { a.getContentResolver().takePersistableUriPermission(u, Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION); } catch (Exception ignored) {}
            NativeDocActivity.open(a, u, nameOf(u));
            return false;
        }
        try {
            byte[] b = read(u);
            JSONObject o = new JSONObject(); o.put("name", nameOf(u)); o.put("b64", Base64.encodeToString(b, Base64.NO_WRAP)); o.put("uri", u.toString()); if (forceWeb) o.put("forceWeb", true);
            try { a.getContentResolver().takePersistableUriPermission(u, Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION); } catch (Exception ignored) {}
            pending = o.toString(); return true;
        } catch (Exception e) { return false; }
    }

    static boolean isOfficeName(String n) {
        String e = n == null ? "" : n.toLowerCase();
        return e.matches(".*\\.(docx|docm|dotx|dotm|doc|dot|rtf|odt|ott|xlsx|xlsm|xltx|xltm|xls|xlt|xlsb|ods|ots|csv|pptx|pptm|ppsx|ppsm|potx|potm|ppt|pps|pot|odp|otp|odg)$");
    }
    boolean nativeEngineWanted() {
        if (!NativeDocActivity.available(a)) return false;
        try { String s = loadSettings(); if (s != null && new JSONObject(s).optString("officeEngine", "native").equals("web")) return false; } catch (Exception ignored) {}
        return true;
    }
    @JavascriptInterface public boolean lokAvailable() { return NativeDocActivity.available(a); }
    @JavascriptInterface public boolean lokOpen(String uri, String name, String b64) {
        try {
            Uri u;
            if (uri != null && uri.length() > 0) u = Uri.parse(uri);
            else {
                java.io.File dir = new java.io.File(a.getCacheDir(), "lok-open/" + Long.toHexString(System.nanoTime())); dir.mkdirs();
                java.io.File f = new java.io.File(dir, name.replaceAll("[\\\\/:*?\"<>|]", "_"));
                java.io.FileOutputStream fo = new java.io.FileOutputStream(f); fo.write(Base64.decode(b64, Base64.DEFAULT)); fo.close();
                u = Uri.fromFile(f);
            }
            final Uri fu = u;
            a.runOnUiThread(() -> NativeDocActivity.open(a, fu, name));
            return true;
        } catch (Exception e) { return false; }
    }

    byte[] read(Uri u) throws Exception {
        InputStream in = a.getContentResolver().openInputStream(u);
        ByteArrayOutputStream bo = new ByteArrayOutputStream(); byte[] buf = new byte[65536]; int n;
        while ((n = in.read(buf)) > 0) bo.write(buf, 0, n); in.close(); return bo.toByteArray();
    }
    String nameOf(Uri u) {
        try (Cursor c = a.getContentResolver().query(u, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
            if (c != null && c.moveToFirst()) { String n = c.getString(0); if (n != null) return n; }
        } catch (Exception ignored) {}
        String p = u.getLastPathSegment(); return p == null ? "tep" : p.substring(p.lastIndexOf('/') + 1);
    }

    // v0.1.8: CPU mà APK này được build cho (theo thư mục lib của gói đã cài), vd "arm64-v8a"
    private String apkAbi() {
        String d = a.getApplicationInfo().nativeLibraryDir; d = d == null ? "" : d;
        if (d.endsWith("arm64")) return "arm64-v8a"; if (d.endsWith("x86_64")) return "x86_64"; if (d.endsWith("arm")) return "armeabi-v7a"; if (d.endsWith("x86")) return "x86";
        return Build.SUPPORTED_ABIS.length > 0 ? Build.SUPPORTED_ABIS[0] : "";
    }
    @JavascriptInterface public String info() {
        try {
            JSONObject o = new JSONObject();
            o.put("platform", "android"); o.put("version", BuildConfig.VERSION_NAME); o.put("sdk", Build.VERSION.SDK_INT); o.put("release", Build.VERSION.RELEASE);
            o.put("device", Build.MANUFACTURER + " " + Build.MODEL); o.put("abi", Build.SUPPORTED_ABIS.length > 0 ? Build.SUPPORTED_ABIS[0] : ""); o.put("edition", BuildConfig.EDITION + "-" + apkAbi()); // v0.1.8
            ActivityManager am = (ActivityManager) a.getSystemService(Context.ACTIVITY_SERVICE); o.put("memClassMB", am.getLargeMemoryClass());
            if (a instanceof MainActivity) {
                o.put("safeInsetsTop", ((MainActivity)a).lastInsetsTop);
                o.put("safeInsetsBottom", ((MainActivity)a).lastInsetsBottom);
            }
            return o.toString();
        } catch (Exception e) { return "{}"; }
    }
    /** Open a file (JSON {name,b64,uri}) or a new document (hash) in a NEW window = new task in Recents (Android 5+ document tasks). */
    @JavascriptInterface public void openWindow(String json, String hash) {
        a.runOnUiThread(() -> {
            try {
                Intent i = new Intent(a, DocActivity.class);
                if (json != null && !json.isEmpty()) { String tok = java.util.UUID.randomUUID().toString(); synchronized (STASH) { STASH.put(tok, json); } i.putExtra("stash", tok); }
                if (hash != null && !hash.isEmpty()) i.putExtra("hash", hash);
                i.addFlags(Intent.FLAG_ACTIVITY_NEW_DOCUMENT | Intent.FLAG_ACTIVITY_MULTIPLE_TASK);
                a.startActivity(i);
            } catch (Exception e) { android.util.Log.e("OOS", "openWindow", e); }
        });
    }
    @JavascriptInterface public boolean canNewWindow() { return true; }
    @JavascriptInterface public String takePendingFile() { String p = pending; pending = null; return p; }

    @JavascriptInterface public String scanFiles() { return doScan(true); }
    // v0.1.8: Feature Check #28 — quét bất đồng bộ (không chặn luồng JS của WebView), KHÔNG tự mở màn hình xin quyền mỗi lần mở app.
    // Trả {"ok":true,"full":<đã có quyền toàn bộ bộ nhớ>,"files":[...]} qua __oosCb.
    @JavascriptInterface public void scanOffice(String id) {
        new Thread(() -> {
            try { JSONObject o = new JSONObject(); o.put("ok", true); o.put("full", hasFullAccess()); o.put("files", new org.json.JSONArray(doScan(false))); a.callback(id, o.toString()); }
            catch (Exception e) { a.callback(id, "{\"ok\":false}"); }
        }).start();
    }
    @JavascriptInterface public void requestFilesAccess() {
        a.runOnUiThread(() -> {
            try {
                if (android.os.Build.VERSION.SDK_INT >= 30) a.startActivity(new Intent(android.provider.Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION, Uri.parse("package:" + a.getPackageName())));
                else a.requestPermissions(new String[]{ android.Manifest.permission.READ_EXTERNAL_STORAGE }, 77);
            } catch (Exception ignored) {}
        });
    }
    private boolean hasFullAccess() {
        return android.os.Build.VERSION.SDK_INT >= 30 ? android.os.Environment.isExternalStorageManager()
            : android.os.Build.VERSION.SDK_INT < 23 || a.checkSelfPermission(android.Manifest.permission.READ_EXTERNAL_STORAGE) == android.content.pm.PackageManager.PERMISSION_GRANTED;
    }
    private String doScan(boolean ask) {
        try {
            org.json.JSONArray arr = new org.json.JSONArray();
            java.util.List<File> roots = new java.util.ArrayList<>();
            try { roots.add(android.os.Environment.getExternalStoragePublicDirectory(android.os.Environment.DIRECTORY_DOCUMENTS)); } catch (Exception ignored) {}
            try { roots.add(android.os.Environment.getExternalStoragePublicDirectory(android.os.Environment.DIRECTORY_DOWNLOADS)); } catch (Exception ignored) {}
            try { roots.add(a.getExternalFilesDir(null)); } catch (Exception ignored) {}
            // v0.1.6: scan the whole shared storage (Feature Check #28) once the user granted access
            boolean full = android.os.Build.VERSION.SDK_INT >= 30 ? android.os.Environment.isExternalStorageManager()
                : android.os.Build.VERSION.SDK_INT < 23 || a.checkSelfPermission(android.Manifest.permission.READ_EXTERNAL_STORAGE) == android.content.pm.PackageManager.PERMISSION_GRANTED;
            if (full) roots.add(0, android.os.Environment.getExternalStorageDirectory());
            else if (ask) a.runOnUiThread(() -> {
                try {
                    if (android.os.Build.VERSION.SDK_INT >= 30) a.startActivity(new Intent(android.provider.Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION, Uri.parse("package:" + a.getPackageName())));
                    else a.requestPermissions(new String[]{ android.Manifest.permission.READ_EXTERNAL_STORAGE }, 77);
                } catch (Exception ignored) {}
            });

            java.util.Set<String> validExts = new java.util.HashSet<>(java.util.Arrays.asList(
                ".docx", ".docm", ".dotx", ".doc",
                ".xlsx", ".xlsm", ".xltx", ".xls", ".csv",
                ".pptx", ".pptm", ".ppsx", ".ppt",
                ".pdf", ".odt", ".ods", ".odp", ".txt", ".md"
            ));

            java.util.Set<String> seen = new java.util.HashSet<>();
            for (File r : roots) {
                if (r != null && r.exists() && r.isDirectory()) {
                    scanDir(r, 0, r.equals(android.os.Environment.getExternalStorageDirectory()) ? 6 : 3, validExts, seen, arr);
                }
            }
            return arr.toString();
        } catch (Exception e) {
            return "[]";
        }
    }

    private void scanDir(File dir, int depth, int maxDepth, java.util.Set<String> exts, java.util.Set<String> seen, org.json.JSONArray arr) {
        try {
            File[] files = dir.listFiles();
            if (files == null) return;
            for (File f : files) {
                if (f.isHidden() || f.getName().startsWith(".") || "Android".equals(f.getName())) continue;
                if (f.isFile()) {
                    String name = f.getName();
                    int dot = name.lastIndexOf('.');
                    if (dot > 0) {
                        String ext = name.substring(dot).toLowerCase(java.util.Locale.US);
                        if (exts.contains(ext) && seen.add(f.getAbsolutePath())) {
                            org.json.JSONObject o = new org.json.JSONObject();
                            o.put("name", name);
                            o.put("path", f.getAbsolutePath());
                            o.put("uri", Uri.fromFile(f).toString());
                            o.put("size", f.length());
                            o.put("mtime", f.lastModified());
                            o.put("dir", dir.getAbsolutePath());
                            o.put("ext", ext.substring(1));
                            String kind = "word";
                            if (ext.startsWith(".xl") || ext.equals(".csv") || ext.equals(".ods")) kind = "excel";
                            else if (ext.startsWith(".pp") || ext.equals(".odp")) kind = "ppt";
                            else if (ext.equals(".pdf")) kind = "pdf";
                            o.put("kind", kind);
                            arr.put(o);
                        }
                    }
                } else if (f.isDirectory() && depth < maxDepth) {
                    scanDir(f, depth + 1, maxDepth, exts, seen, arr);
                }
            }
        } catch (Exception ignored) {}
    }

    @JavascriptInterface public void saveFile(String name, String mime, String b64, String id) {
        saveBytes = Base64.decode(b64, Base64.DEFAULT); saveName = name; saveId = id;
        Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT); i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType(mime == null || mime.isEmpty() ? "application/octet-stream" : mime); i.putExtra(Intent.EXTRA_TITLE, name);
        a.runOnUiThread(() -> { try { a.startActivityForResult(i, MainActivity.REQ_CREATE); } catch (Exception e) { a.callback(saveId, "{\"ok\":false,\"error\":\"no picker\"}"); } });
    }
    void onCreateResult(Uri u) {
        String id = saveId; saveId = null;
        if (u == null) { saveBytes = null; a.callback(id, "{\"ok\":false}"); return; }
        try {
            OutputStream os = a.getContentResolver().openOutputStream(u, "wt"); os.write(saveBytes); os.close(); saveBytes = null;
            try { a.getContentResolver().takePersistableUriPermission(u, Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION); } catch (Exception ignored) {}
            JSONObject o = new JSONObject(); o.put("ok", true); o.put("name", nameOf(u)); o.put("uri", u.toString()); a.callback(id, o.toString());
        } catch (Exception e) { a.callback(id, "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
    }
    @JavascriptInterface public void writeUri(String uri, String b64, String id) {
        new Thread(() -> {
            try { Uri u = Uri.parse(uri); OutputStream os = a.getContentResolver().openOutputStream(u, "wt"); os.write(Base64.decode(b64, Base64.DEFAULT)); os.close();
                JSONObject o = new JSONObject(); o.put("ok", true); o.put("name", nameOf(u)); a.callback(id, o.toString()); }
            catch (Exception e) { a.callback(id, "{\"ok\":false}"); }
        }).start();
    }
    @JavascriptInterface public void readUri(String uri, String id) {
        new Thread(() -> {
            try { Uri u = Uri.parse(uri); byte[] b = read(u); JSONObject o = new JSONObject(); o.put("ok", true); o.put("name", nameOf(u)); o.put("b64", Base64.encodeToString(b, Base64.NO_WRAP)); a.callback(id, o.toString()); }
            catch (Exception e) { a.callback(id, "{\"ok\":false}"); }
        }).start();
    }
    @JavascriptInterface public void print(String title) {
        a.runOnUiThread(() -> {
            PrintManager pm = (PrintManager) a.getSystemService(Context.PRINT_SERVICE);
            String t = title == null || title.isEmpty() ? "Office Open Source" : title;
            pm.print(t, a.web.createPrintDocumentAdapter(t), new PrintAttributes.Builder().build());
        });
    }
    @JavascriptInterface public void share(String name, String mime, String b64) {
        try {
            File dir = new File(a.getCacheDir(), "share"); dir.mkdirs();
            String safe = name.replaceAll("[\\\\/:*?\"<>|]", "_"); File f = new File(dir, safe);
            FileOutputStream fo = new FileOutputStream(f); fo.write(Base64.decode(b64, Base64.DEFAULT)); fo.close();
            Uri u = Uri.parse("content://" + ShareProvider.AUTH + "/" + Uri.encode(safe));
            Intent i = new Intent(Intent.ACTION_SEND); i.setType(mime); i.putExtra(Intent.EXTRA_STREAM, u); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            a.runOnUiThread(() -> a.startActivity(Intent.createChooser(i, "Chia sẻ " + name)));
        } catch (Exception ignored) {}
    }
    @JavascriptInterface public void openUrl(String url) { a.runOnUiThread(() -> { try { a.startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(url))); } catch (Exception ignored) {} }); }
    // ---------------- v0.0.7: A.I web screen, clipboard, HTTP GET (updater listings), download + install update APK
    @JavascriptInterface public void aiWeb(String url, String title) {
        a.runOnUiThread(() -> { try { Intent i = new Intent(a, AiWebActivity.class); i.putExtra("url", url); i.putExtra("title", title); a.startActivityForResult(i, MainActivity.REQ_AIWEB); } catch (Exception e) { android.util.Log.e("OOS", "aiWeb", e); } });
    }
    @JavascriptInterface public String clipText() {
        final String[] out = {""}; final java.util.concurrent.CountDownLatch l = new java.util.concurrent.CountDownLatch(1);
        a.runOnUiThread(() -> { try { android.content.ClipboardManager cm = (android.content.ClipboardManager) a.getSystemService(Context.CLIPBOARD_SERVICE); android.content.ClipData c = cm.getPrimaryClip(); if (c != null && c.getItemCount() > 0) { CharSequence t = c.getItemAt(0).coerceToText(a); out[0] = t == null ? "" : t.toString(); } } catch (Exception ignored) {} l.countDown(); });
        try { l.await(2, java.util.concurrent.TimeUnit.SECONDS); } catch (InterruptedException ignored) {}
        return out[0];
    }
    @JavascriptInterface public void clipSet(String t) { a.runOnUiThread(() -> { try { android.content.ClipboardManager cm = (android.content.ClipboardManager) a.getSystemService(Context.CLIPBOARD_SERVICE); cm.setPrimaryClip(android.content.ClipData.newPlainText("Office Open Source", t)); } catch (Exception ignored) {} }); }
    static java.net.HttpURLConnection open(String url) throws Exception {
        java.net.URL u = new java.net.URL(url); java.net.HttpURLConnection c = null;
        for (int hop = 0; hop < 8; hop++) {
            c = (java.net.HttpURLConnection) u.openConnection(); c.setInstanceFollowRedirects(false); c.setConnectTimeout(20000); c.setReadTimeout(60000);
            c.setRequestProperty("User-Agent", "Mozilla/5.0 (Linux; Android " + Build.VERSION.RELEASE + ") OfficeOpenSource/" + BuildConfig.VERSION_NAME);
            int code = c.getResponseCode();
            if (code >= 300 && code < 400 && c.getHeaderField("Location") != null) { u = new java.net.URL(u, c.getHeaderField("Location")); if (!"https".equals(u.getProtocol())) throw new Exception("blocked redirect"); c.disconnect(); continue; }
            return c;
        }
        return c;
    }
    @JavascriptInterface public void httpGet(String url, String id) {
        new Thread(() -> {
            try { if (!url.startsWith("https://")) throw new Exception("blocked"); java.net.HttpURLConnection c = open(url); InputStream in = c.getInputStream(); ByteArrayOutputStream bo = new ByteArrayOutputStream(); byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) bo.write(buf, 0, n); in.close();
                JSONObject o = new JSONObject(); o.put("ok", true); o.put("body", new String(bo.toByteArray(), "UTF-8")); a.callback(id, o.toString()); }
            catch (Exception e) { a.callback(id, "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
        }).start();
    }
    // ---------------- v0.0.8: generic HTTPS for cloud connectors (method, headers, base64 body/response; follows redirects)
    @JavascriptInterface public void http(String json, String id) {
        new Thread(() -> {
            try {
                JSONObject q = new JSONObject(json); String url = q.getString("url"); String method = q.optString("method", "GET").toUpperCase();
                if (!url.startsWith("https://")) throw new Exception("blocked");
                JSONObject hd = q.optJSONObject("headers"); String b64 = q.optString("b64", "");
                byte[] body = b64.isEmpty() ? null : Base64.decode(b64, Base64.NO_WRAP);
                java.net.URL u = new java.net.URL(url); java.net.HttpURLConnection c = null; int code = 0;
                for (int hop = 0; hop < 8; hop++) {
                    c = (java.net.HttpURLConnection) u.openConnection(); c.setInstanceFollowRedirects(false); c.setConnectTimeout(30000); c.setReadTimeout(300000);
                    // HttpURLConnection has no PATCH: Google APIs accept POST + X-HTTP-Method-Override
                    if ("PATCH".equals(method)) { c.setRequestMethod("POST"); c.setRequestProperty("X-HTTP-Method-Override", "PATCH"); } else c.setRequestMethod(hop > 0 && code != 307 && code != 308 ? "GET" : method);
                    c.setRequestProperty("User-Agent", "OfficeOpenSource/" + BuildConfig.VERSION_NAME + " (Android " + Build.VERSION.RELEASE + ")");
                    if (hd != null) { boolean sameHost = u.getHost().equals(new java.net.URL(url).getHost()); java.util.Iterator<String> it = hd.keys(); while (it.hasNext()) { String k = it.next(); if (hop > 0 && (!sameHost && "Authorization".equalsIgnoreCase(k) || "Content-Type".equalsIgnoreCase(k) && code != 307 && code != 308)) continue; c.setRequestProperty(k, hd.getString(k)); } }
                    boolean send = hop == 0 && (body != null || "POST".equals(method) || "PUT".equals(method) || "PATCH".equals(method));
                    if (send) { c.setDoOutput(true); byte[] b = body == null ? new byte[0] : body; c.setFixedLengthStreamingMode(b.length); OutputStream os = c.getOutputStream(); os.write(b); os.close(); }
                    code = c.getResponseCode();
                    if (code >= 300 && code < 400 && c.getHeaderField("Location") != null) { u = new java.net.URL(u, c.getHeaderField("Location")); if (!"https".equals(u.getProtocol())) throw new Exception("blocked redirect"); c.disconnect(); continue; }
                    break;
                }
                InputStream in = code >= 400 ? c.getErrorStream() : c.getInputStream();
                ByteArrayOutputStream bo = new ByteArrayOutputStream(); if (in != null) { byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) bo.write(buf, 0, n); in.close(); }
                JSONObject o = new JSONObject(); o.put("status", code); o.put("type", c.getContentType() == null ? "" : c.getContentType()); o.put("b64", Base64.encodeToString(bo.toByteArray(), Base64.NO_WRAP));
                a.callback(id, o.toString());
            } catch (Exception e) { a.callback(id, "{\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
        }).start();
    }
    // ---------------- v0.0.8: OAuth sign-in in the system browser, redirect caught on http://127.0.0.1:<port>/ (RFC 8252 loopback, like rclone)
    static java.net.ServerSocket oauthSock;
    @JavascriptInterface public void oauth(String url, String port, String id) {
        new Thread(() -> {
            java.net.ServerSocket ss = null;
            try {
                if (!url.startsWith("https://")) throw new Exception("blocked");
                synchronized (Bridge.class) { if (oauthSock != null) try { oauthSock.close(); } catch (Exception ignored) {} ss = new java.net.ServerSocket(); ss.setReuseAddress(true); ss.bind(new java.net.InetSocketAddress(java.net.InetAddress.getByName("127.0.0.1"), Integer.parseInt(port))); oauthSock = ss; }
                ss.setSoTimeout(600000);
                a.runOnUiThread(() -> { try { Intent i = new Intent(Intent.ACTION_VIEW, Uri.parse(url)); i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK); a.startActivity(i); } catch (Exception ignored) {} });
                while (true) {
                    java.net.Socket s = ss.accept();
                    try {
                        s.setSoTimeout(10000); InputStream in = s.getInputStream(); byte[] buf = new byte[16384]; int n = in.read(buf);
                        String line = new String(buf, 0, Math.max(0, n), "US-ASCII").split("\n")[0]; String[] parts = line.split(" "); String path = parts.length > 1 ? parts[1] : "/";
                        boolean hit = path.contains("code=") || path.contains("error=");
                        String html = hit ? "<!doctype html><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Office Open Source</title><body style='font-family:sans-serif;text-align:center;padding:48px'><h2>" + (path.contains("code=") ? "Đăng nhập thành công" : "Đăng nhập bị hủy") + "</h2><p>Hãy quay lại ứng dụng Office Open Source.</p>" : "";
                        byte[] b = html.getBytes("UTF-8"); OutputStream os = s.getOutputStream();
                        os.write(((hit ? "HTTP/1.1 200 OK" : "HTTP/1.1 404 Not Found") + "\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: " + b.length + "\r\nConnection: close\r\n\r\n").getBytes("US-ASCII")); os.write(b); os.flush(); s.close();
                        if (hit) {
                            JSONObject o = new JSONObject(); o.put("ok", true); o.put("query", path); a.callback(id, o.toString());
                            a.runOnUiThread(() -> { try { Intent back = new Intent(a, a.getClass()); back.addFlags(Intent.FLAG_ACTIVITY_REORDER_TO_FRONT | Intent.FLAG_ACTIVITY_SINGLE_TOP); a.startActivity(back); } catch (Exception ignored) {} });
                            break;
                        }
                    } catch (Exception ignored) { try { s.close(); } catch (Exception e2) {} }
                }
            } catch (Exception e) { a.callback(id, "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
            finally { try { if (ss != null) ss.close(); } catch (Exception ignored) {} }
        }).start();
    }
    @JavascriptInterface public void downloadInstall(String url, String name, String id) {
        new Thread(() -> {
            try {
                if (!url.startsWith("https://")) throw new Exception("blocked");
                String safe = new File(name == null || name.isEmpty() ? "OfficeOpenSource-update.apk" : name).getName();
                File dir = Build.VERSION.SDK_INT >= 24 ? new File(a.getCacheDir(), "share") : new File(a.getExternalCacheDir(), "update"); dir.mkdirs();
                File f = new File(dir, safe);
                java.net.HttpURLConnection c = open(url);
                String ct = c.getContentType(); if (ct != null && ct.startsWith("text/html")) throw new Exception("Máy chủ trả về trang web thay vì tệp APK");
                InputStream in = c.getInputStream(); FileOutputStream fo = new FileOutputStream(f); byte[] buf = new byte[65536]; int n; while ((n = in.read(buf)) > 0) fo.write(buf, 0, n); fo.close(); in.close();
                f.setReadable(true, false);
                a.runOnUiThread(() -> {
                    try {
                        if (Build.VERSION.SDK_INT >= 26 && !a.getPackageManager().canRequestPackageInstalls()) {
                            a.startActivity(new Intent(android.provider.Settings.ACTION_MANAGE_UNKNOWN_APP_SOURCES, Uri.parse("package:" + a.getPackageName())));
                            android.widget.Toast.makeText(a, "Cho phép \"Cài ứng dụng không rõ nguồn\" rồi bấm Cập nhật lại", android.widget.Toast.LENGTH_LONG).show();
                            a.callback(id, "{\"ok\":true,\"needPermission\":true}"); return;
                        }
                        Intent i = new Intent(Intent.ACTION_VIEW);
                        Uri uri = Build.VERSION.SDK_INT >= 24 ? Uri.parse("content://" + ShareProvider.AUTH + "/" + Uri.encode(safe)) : Uri.fromFile(f);
                        i.setDataAndType(uri, "application/vnd.android.package-archive"); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_ACTIVITY_NEW_TASK);
                        a.startActivity(i); a.callback(id, "{\"ok\":true}");
                    } catch (Exception e) { a.callback(id, "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
                });
            } catch (Exception e) { a.callback(id, "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}"); }
        }).start();
    }
    @JavascriptInterface public void saveSettings(String json) {
        try {
            android.content.SharedPreferences sp = a.getSharedPreferences("oos_settings", Context.MODE_PRIVATE);
            sp.edit().putString("data", json).apply();
        } catch (Exception ignored) {}
    }
    @JavascriptInterface public String loadSettings() {
        try {
            android.content.SharedPreferences sp = a.getSharedPreferences("oos_settings", Context.MODE_PRIVATE);
            return sp.getString("data", null);
        } catch (Exception e) { return null; }
    }
    @JavascriptInterface public void setTitle(String t) { a.runOnUiThread(() -> { try { a.setTaskDescription(new ActivityManager.TaskDescription(t)); } catch (Exception ignored) {} }); }
    @JavascriptInterface public void exit() { a.runOnUiThread(a::finish); }

    // =========================================================================
    // Android DeX Mode, External Displays & Phone-to-PC Web Link Engine (v0.1.3)
    // =========================================================================

    @JavascriptInterface public boolean isSamsungDeX() {
        try {
            android.content.res.Configuration config = a.getResources().getConfiguration();
            Class<?> configClass = config.getClass();
            java.lang.reflect.Field field = configClass.getField("SEM_DESKTOP_MODE_ENABLED");
            int enabledVal = field.getInt(null);
            java.lang.reflect.Field curModeField = configClass.getField("semDesktopModeEnabled");
            int curMode = curModeField.getInt(config);
            return curMode == enabledVal;
        } catch (Exception e) {
            return false;
        }
    }

    @JavascriptInterface public String getDisplayInfo() {
        try {
            JSONObject res = new JSONObject();
            res.put("isDex", isSamsungDeX());
            android.hardware.display.DisplayManager dm = (android.hardware.display.DisplayManager) a.getSystemService(Context.DISPLAY_SERVICE);
            android.view.Display[] displays = dm.getDisplays();
            org.json.JSONArray arr = new org.json.JSONArray();
            for (android.view.Display d : displays) {
                JSONObject dObj = new JSONObject();
                dObj.put("id", d.getDisplayId());
                dObj.put("name", d.getName());
                android.util.DisplayMetrics metrics = new android.util.DisplayMetrics();
                d.getRealMetrics(metrics);
                dObj.put("width", metrics.widthPixels);
                dObj.put("height", metrics.heightPixels);
                dObj.put("densityDpi", metrics.densityDpi);
                dObj.put("refreshRate", (double) d.getRefreshRate());
                arr.put(dObj);
            }
            res.put("displays", arr);
            return res.toString();
        } catch (Exception e) {
            return "{\"isDex\":false,\"displays\":[]}";
        }
    }

    @JavascriptInterface public String getDeviceNetworkInfo() {
        try {
            JSONObject res = new JSONObject();
            String ipAddress = "127.0.0.1";
            java.util.Enumeration<java.net.NetworkInterface> interfaces = java.net.NetworkInterface.getNetworkInterfaces();
            while (interfaces.hasMoreElements()) {
                java.net.NetworkInterface ni = interfaces.nextElement();
                if (ni.isLoopback() || !ni.isUp()) continue;
                java.util.Enumeration<java.net.InetAddress> addresses = ni.getInetAddresses();
                while (addresses.hasMoreElements()) {
                    java.net.InetAddress addr = addresses.nextElement();
                    if (!addr.isLoopbackAddress() && addr instanceof java.net.Inet4Address) {
                        ipAddress = addr.getHostAddress();
                        break;
                    }
                }
            }
            res.put("ip", ipAddress);
            return res.toString();
        } catch (Exception e) {
            return "{\"ip\":\"127.0.0.1\"}";
        }
    }

    @JavascriptInterface public void setKeepScreenOn(boolean keep) {
        a.runOnUiThread(() -> {
            try {
                if (keep) a.getWindow().addFlags(android.view.WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
                else a.getWindow().clearFlags(android.view.WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
            } catch (Exception ignored) {}
        });
    }

    @JavascriptInterface public void openSystemCastSettings() {
        a.runOnUiThread(() -> {
            try {
                Intent i = new Intent(android.provider.Settings.ACTION_CAST_SETTINGS);
                i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                a.startActivity(i);
            } catch (Exception e) {
                try {
                    Intent i = new Intent("android.settings.WIFI_DISPLAY_SETTINGS");
                    i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                    a.startActivity(i);
                } catch (Exception e2) {
                    try {
                        Intent i = new Intent(android.provider.Settings.ACTION_DISPLAY_SETTINGS);
                        i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                        a.startActivity(i);
                    } catch (Exception ignored) {}
                }
            }
        });
    }

    // v0.1.6: list launchable apps with MAIN/LAUNCHER (works with Android 11+ package visibility via <queries>)
    @JavascriptInterface public String getInstalledApps() {
        try {
            android.content.pm.PackageManager pm = a.getPackageManager();
            Intent q = new Intent(Intent.ACTION_MAIN); q.addCategory(Intent.CATEGORY_LAUNCHER);
            java.util.List<android.content.pm.ResolveInfo> list = pm.queryIntentActivities(q, 0);
            java.util.Collections.sort(list, new android.content.pm.ResolveInfo.DisplayNameComparator(pm));
            org.json.JSONArray arr = new org.json.JSONArray(); java.util.HashSet<String> seen = new java.util.HashSet<>();
            for (android.content.pm.ResolveInfo ri : list) {
                String pkg = ri.activityInfo.packageName;
                if (pkg.equals(a.getPackageName()) || !seen.add(pkg)) continue;
                JSONObject o = new JSONObject();
                o.put("packageName", pkg);
                o.put("name", ri.loadLabel(pm).toString());
                o.put("isSystem", (ri.activityInfo.applicationInfo.flags & android.content.pm.ApplicationInfo.FLAG_SYSTEM) != 0);
                arr.put(o);
            }
            return arr.toString();
        } catch (Exception e) {
            return "[]";
        }
    }
    static final java.util.HashMap<String, String> ICONS = new java.util.HashMap<>();
    @JavascriptInterface public String getAppIcon(String pkg) {
        synchronized (ICONS) { String c = ICONS.get(pkg); if (c != null) return c; }
        try {
            android.graphics.drawable.Drawable d = a.getPackageManager().getApplicationIcon(pkg);
            int sz = 96; android.graphics.Bitmap bm = android.graphics.Bitmap.createBitmap(sz, sz, android.graphics.Bitmap.Config.ARGB_8888);
            android.graphics.Canvas cv = new android.graphics.Canvas(bm); d.setBounds(0, 0, sz, sz); d.draw(cv);
            java.io.ByteArrayOutputStream bo = new java.io.ByteArrayOutputStream(); bm.compress(android.graphics.Bitmap.CompressFormat.PNG, 100, bo);
            String r = "data:image/png;base64," + Base64.encodeToString(bo.toByteArray(), Base64.NO_WRAP);
            synchronized (ICONS) { ICONS.put(pkg, r); } return r;
        } catch (Exception e) { return ""; }
    }

    @JavascriptInterface public boolean launchApp(String packageName) {
        return launchAppWithOptions(packageName, -1, -1, -1, -1, -1);
    }

    @JavascriptInterface public boolean launchAppWithOptions(String packageName, int x, int y, int w, int h, int displayId) {
        if (packageName == null || packageName.trim().isEmpty()) return false;
        packageName = packageName.trim();
        try {
            android.content.pm.PackageManager pm = a.getPackageManager();
            Intent i = pm.getLaunchIntentForPackage(packageName);
            if (i == null) return false;
            i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            try {
                i.addFlags(Intent.FLAG_ACTIVITY_MULTIPLE_TASK);
            } catch (Throwable ignored) {}

            android.os.Bundle optionsBundle = null;
            if (android.os.Build.VERSION.SDK_INT >= 24) {
                android.app.ActivityOptions options = android.app.ActivityOptions.makeBasic();
                if (w > 0 && h > 0 && x >= 0 && y >= 0) {
                    // freeform window (DeX / desktop mode). Split-screen (x == 0 or right half, full height) also asks for the adjacent pane.
                    options.setLaunchBounds(new android.graphics.Rect(x, y, x + w, y + h));
                    android.util.DisplayMetrics dm = a.getResources().getDisplayMetrics();
                    if (h >= dm.heightPixels * 0.9 && w <= dm.widthPixels * 0.55) try { i.addFlags(Intent.FLAG_ACTIVITY_LAUNCH_ADJACENT); } catch (Throwable ignored) {}
                }
                if (displayId >= 0 && android.os.Build.VERSION.SDK_INT >= 26) {
                    try {
                        options.setLaunchDisplayId(displayId);
                    } catch (Throwable ignored) {}
                }
                optionsBundle = options.toBundle();
            }
            if (optionsBundle != null) {
                a.startActivity(i, optionsBundle);
            } else {
                a.startActivity(i);
            }
            return true;
        } catch (Throwable t) {
            try {
                android.content.pm.PackageManager pm = a.getPackageManager();
                Intent i = pm.getLaunchIntentForPackage(packageName);
                if (i != null) {
                    i.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                    a.startActivity(i);
                    return true;
                }
            } catch (Throwable ignored) {}
        }
        return false;
    }

    // v0.1.6: AirDrop / QuickShare over Wi-Fi
    @JavascriptInterface public String lanShare(String name, String b64) {
        try { return LanShare.share(a, name, Base64.decode(b64, Base64.DEFAULT)); } catch (Exception e) { return "{\"ok\":false}"; }
    }
    @JavascriptInterface public String lanReceived() {
        org.json.JSONArray arr = new org.json.JSONArray();
        synchronized (LanShare.received) { for (String p : LanShare.received) arr.put(Uri.fromFile(new java.io.File(p)).toString()); LanShare.received.clear(); }
        return arr.toString();
    }
    // v0.1.6: DeX Quick Settings helpers
    @JavascriptInterface public void openSystemSettings(String what) {
        a.runOnUiThread(() -> {
            String act = "wifi".equals(what) ? android.provider.Settings.ACTION_WIFI_SETTINGS : "bluetooth".equals(what) ? android.provider.Settings.ACTION_BLUETOOTH_SETTINGS
                : "cast".equals(what) ? "android.settings.CAST_SETTINGS" : "display".equals(what) ? android.provider.Settings.ACTION_DISPLAY_SETTINGS : android.provider.Settings.ACTION_SETTINGS;
            try { a.startActivity(new Intent(act).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)); } catch (Exception e) { try { a.startActivity(new Intent(android.provider.Settings.ACTION_SETTINGS).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)); } catch (Exception ignored) {} }
        });
    }
    @JavascriptInterface public void setVolume(int pct) {
        try { android.media.AudioManager am = (android.media.AudioManager) a.getSystemService(Context.AUDIO_SERVICE); int mx = am.getStreamMaxVolume(android.media.AudioManager.STREAM_MUSIC); am.setStreamVolume(android.media.AudioManager.STREAM_MUSIC, Math.round(mx * Math.max(0, Math.min(100, pct)) / 100f), 0); } catch (Exception ignored) {}
    }
    @JavascriptInterface public void setBrightness(float v) {
        a.runOnUiThread(() -> { try { android.view.WindowManager.LayoutParams lp = a.getWindow().getAttributes(); lp.screenBrightness = Math.max(0.02f, Math.min(1f, v)); a.getWindow().setAttributes(lp); } catch (Exception ignored) {} });
    }
    @JavascriptInterface public int getBattery() {
        try { Intent b = a.registerReceiver(null, new android.content.IntentFilter(Intent.ACTION_BATTERY_CHANGED)); int l = b.getIntExtra("level", -1), sc = b.getIntExtra("scale", 100); return l < 0 ? 100 : Math.round(l * 100f / sc); } catch (Exception e) { return 100; }
    }

    static java.net.ServerSocket webLinkSock;
    static boolean webLinkRunning = false;
    static int webLinkPort = 8080;

    @JavascriptInterface public String startWebLink(int port) {
        try {
            if (webLinkRunning && webLinkSock != null) {
                return "{\"ok\":true,\"port\":" + webLinkPort + ",\"running\":true}";
            }
            webLinkPort = port > 0 ? port : 8080;
            java.net.ServerSocket ss = new java.net.ServerSocket();
            ss.setReuseAddress(true);
            ss.bind(new java.net.InetSocketAddress("0.0.0.0", webLinkPort));
            webLinkSock = ss;
            webLinkRunning = true;

            new Thread(() -> {
                while (webLinkRunning && !ss.isClosed()) {
                    try {
                        java.net.Socket s = ss.accept();
                        new Thread(() -> handleWebLinkClient(s)).start();
                    } catch (Exception ignored) {}
                }
            }).start();

            JSONObject res = new JSONObject();
            res.put("ok", true);
            res.put("port", webLinkPort);
            res.put("running", true);
            return res.toString();
        } catch (Exception e) {
            return "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}";
        }
    }

    @JavascriptInterface public String stopWebLink() {
        try {
            webLinkRunning = false;
            if (webLinkSock != null) {
                webLinkSock.close();
                webLinkSock = null;
            }
            return "{\"ok\":true,\"running\":false}";
        } catch (Exception e) {
            return "{\"ok\":false,\"error\":" + JSONObject.quote(String.valueOf(e.getMessage())) + "}";
        }
    }

    @JavascriptInterface public String getWebLinkStatus() {
        try {
            JSONObject res = new JSONObject();
            res.put("running", webLinkRunning);
            res.put("port", webLinkPort);
            return res.toString();
        } catch (Exception e) {
            return "{\"running\":false,\"port\":8080}";
        }
    }

    void handleWebLinkClient(java.net.Socket s) {
        try {
            s.setSoTimeout(15000);
            InputStream in = s.getInputStream();
            OutputStream os = s.getOutputStream();
            byte[] buf = new byte[8192];
            int n = in.read(buf);
            if (n <= 0) { s.close(); return; }
            String reqLine = new String(buf, 0, n, "US-ASCII").split("\r?\n")[0];
            String[] parts = reqLine.split(" ");
            String path = parts.length > 1 ? parts[1] : "/";
            if (path.contains("?")) path = path.substring(0, path.indexOf('?'));
            if (path.equals("/") || path.isEmpty()) path = "/index.html";

            String assetPath = "www" + path;
            try {
                InputStream assetIn = a.getAssets().open(assetPath);
                ByteArrayOutputStream bo = new ByteArrayOutputStream();
                byte[] fbuf = new byte[65536];
                int fn;
                while ((fn = assetIn.read(fbuf)) > 0) bo.write(fbuf, 0, fn);
                assetIn.close();
                byte[] content = bo.toByteArray();

                String ext = path.substring(path.lastIndexOf('.') + 1).toLowerCase();
                String mime = MainActivity.MIME.containsKey(ext) ? MainActivity.MIME.get(ext) : "application/octet-stream";

                String header = "HTTP/1.1 200 OK\r\n" +
                    "Content-Type: " + mime + (mime.startsWith("text/") || mime.contains("javascript") || mime.contains("json") ? "; charset=utf-8" : "") + "\r\n" +
                    "Content-Length: " + content.length + "\r\n" +
                    "Access-Control-Allow-Origin: *\r\n" +
                    "Connection: close\r\n\r\n";
                os.write(header.getBytes("US-ASCII"));
                os.write(content);
                os.flush();
            } catch (Exception notFound) {
                String nf = "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n";
                os.write(nf.getBytes("US-ASCII"));
                os.flush();
            }
            s.close();
        } catch (Exception ignored) {
            try { s.close(); } catch (Exception e2) {}
        }
    }
}
