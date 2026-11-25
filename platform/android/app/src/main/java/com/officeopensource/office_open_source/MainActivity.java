package com.officeopensource.office_open_source;

import android.app.Activity;
import android.content.Intent;
import android.content.pm.ApplicationInfo;
import android.net.Uri;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.webkit.ConsoleMessage;
import android.webkit.PermissionRequest;
import android.content.pm.PackageManager;
import android.Manifest;
import java.util.ArrayList;
import java.util.List;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/** Office Open Source v0.0.8 - native WebView shell (Android 6.0 / API 23+). Serves the bundled web app from assets at https://oos.app/ */
public class MainActivity extends Activity {
    static final String TAG = "OOS";
    static final String ORIGIN = "https://oos.app/";
    static final int REQ_CHOOSER = 41, REQ_CREATE = 42, REQ_MEDIA = 43, REQ_AIWEB = 44;
    PermissionRequest pendingMedia;
    WebView web;
    Bridge bridge;
    ValueCallback<Uri[]> chooserCb;

    int lastInsetsTop = 24, lastInsetsBottom = 48, lastInsetsLeft = 0, lastInsetsRight = 0;

    void applySafeInsets() {
        if (web == null) return;
        final String js = "try { " +
            "document.documentElement.style.setProperty('--safe-top', '" + lastInsetsTop + "px');" +
            "document.documentElement.style.setProperty('--safe-bottom', '" + lastInsetsBottom + "px');" +
            "document.documentElement.style.setProperty('--safe-left', '" + lastInsetsLeft + "px');" +
            "document.documentElement.style.setProperty('--safe-right', '" + lastInsetsRight + "px');" +
            "} catch(e) {}";
        web.post(() -> {
            if (android.os.Build.VERSION.SDK_INT >= 19) web.evaluateJavascript(js, null);
            else web.loadUrl("javascript:" + js);
        });
    }

    @Override protected void onCreate(Bundle b) {
        super.onCreate(b);

        try {
            float density = getResources().getDisplayMetrics().density;
            if (density <= 0) density = 1.0f;
            int sbRes = getResources().getIdentifier("status_bar_height", "dimen", "android");
            if (sbRes > 0) {
                lastInsetsTop = (int) Math.ceil(getResources().getDimensionPixelSize(sbRes) / density);
            }
            int nbRes = getResources().getIdentifier("navigation_bar_height", "dimen", "android");
            if (nbRes > 0) {
                lastInsetsBottom = (int) Math.ceil(getResources().getDimensionPixelSize(nbRes) / density);
            }
        } catch (Exception ignored) {}

        web = new WebView(this);
        web.setFitsSystemWindows(true);
        if (android.os.Build.VERSION.SDK_INT >= 20) {
            web.setOnApplyWindowInsetsListener((v, insets) -> {
                float density = getResources().getDisplayMetrics().density;
                if (density <= 0) density = 1.0f;
                int top = insets.getSystemWindowInsetTop();
                int bottom = insets.getSystemWindowInsetBottom();
                int left = insets.getSystemWindowInsetLeft();
                int right = insets.getSystemWindowInsetRight();
                if (top > 0) lastInsetsTop = (int) Math.ceil(top / density);
                if (bottom > 0) lastInsetsBottom = (int) Math.ceil(bottom / density);
                if (left > 0) lastInsetsLeft = (int) Math.ceil(left / density);
                if (right > 0) lastInsetsRight = (int) Math.ceil(right / density);
                applySafeInsets();
                return insets.consumeSystemWindowInsets();
            });
        }

        if (android.os.Build.VERSION.SDK_INT >= 21) {
            try {
                getWindow().setStatusBarColor(0xFFF3F3F3);
                getWindow().setNavigationBarColor(0xFFF3F3F3);
            } catch (Exception ignored) {}
        }
        if (android.os.Build.VERSION.SDK_INT >= 23) {
            try {
                View decor = getWindow().getDecorView();
                int flags = decor.getSystemUiVisibility();
                flags |= View.SYSTEM_UI_FLAG_LIGHT_STATUS_BAR;
                decor.setSystemUiVisibility(flags);
            } catch (Exception ignored) {}
        }
        if (android.os.Build.VERSION.SDK_INT >= 26) {
            try {
                View decor = getWindow().getDecorView();
                int flags = decor.getSystemUiVisibility();
                flags |= View.SYSTEM_UI_FLAG_LIGHT_NAVIGATION_BAR;
                decor.setSystemUiVisibility(flags);
            } catch (Exception ignored) {}
        }

        setContentView(web);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setAllowFileAccess(false);
        s.setAllowContentAccess(true);
        s.setLoadWithOverviewMode(true);
        s.setUseWideViewPort(true);
        s.setBuiltInZoomControls(false);
        s.setSupportZoom(false);
        s.setTextZoom(100);
        s.setMediaPlaybackRequiresUserGesture(false);
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_ALWAYS_ALLOW); // allow local LLM servers (Ollama/LM Studio) from the app
        s.setCacheMode(WebSettings.LOAD_DEFAULT);
        s.setUserAgentString(s.getUserAgentString() + " OfficeOpenSource/0.1.7 HybridNative");
        web.setLayerType(View.LAYER_TYPE_HARDWARE, null);
        web.setOverScrollMode(View.OVER_SCROLL_NEVER);
        web.setVerticalScrollBarEnabled(false);
        web.setHorizontalScrollBarEnabled(false);
        bridge = new Bridge(this);
        web.addJavascriptInterface(bridge, "OOSNative");
        web.setWebViewClient(new WebViewClient() {
            @Override public WebResourceResponse shouldInterceptRequest(WebView v, WebResourceRequest r) { return serve(r.getUrl()); }
            @SuppressWarnings("deprecation") @Override public WebResourceResponse shouldInterceptRequest(WebView v, String url) { return serve(Uri.parse(url)); }
            @Override public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest r) { return external(r.getUrl()); }
            @SuppressWarnings("deprecation") @Override public boolean shouldOverrideUrlLoading(WebView v, String url) { return external(Uri.parse(url)); }
            @Override public void onPageFinished(WebView v, String url) { super.onPageFinished(v, url); applySafeInsets(); }
        });
        web.setWebChromeClient(new WebChromeClient() {
            @Override public boolean onConsoleMessage(ConsoleMessage m) {
                String t = "[js:" + m.messageLevel() + "] " + m.message() + " @" + m.sourceId() + ":" + m.lineNumber();
                if (m.messageLevel() == ConsoleMessage.MessageLevel.ERROR) Log.e(TAG, t); else Log.i(TAG, t);
                return true;
            }
            // v0.0.6: microphone / camera for PowerPoint Record tab (narration, cameo) — maps WebView media requests to Android runtime permissions
            @Override public void onPermissionRequest(final PermissionRequest r) {
                runOnUiThread(() -> {
                    List<String> need = new ArrayList<>();
                    for (String res : r.getResources()) {
                        if (PermissionRequest.RESOURCE_AUDIO_CAPTURE.equals(res) && checkSelfPermission(Manifest.permission.RECORD_AUDIO) != PackageManager.PERMISSION_GRANTED) need.add(Manifest.permission.RECORD_AUDIO);
                        if (PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(res) && checkSelfPermission(Manifest.permission.CAMERA) != PackageManager.PERMISSION_GRANTED) need.add(Manifest.permission.CAMERA);
                    }
                    if (need.isEmpty()) { grantMedia(r); return; }
                    if (pendingMedia != null) pendingMedia.deny();
                    pendingMedia = r;
                    requestPermissions(need.toArray(new String[0]), REQ_MEDIA);
                });
            }
            @Override public boolean onShowFileChooser(WebView v, ValueCallback<Uri[]> cb, FileChooserParams p) {
                if (chooserCb != null) chooserCb.onReceiveValue(null);
                chooserCb = cb;
                Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
                i.addCategory(Intent.CATEGORY_OPENABLE);
                String[] acc = p.getAcceptTypes();
                String mime = (acc != null && acc.length > 0 && acc[0] != null && acc[0].contains("/")) ? acc[0] : "*/*";
                i.setType(mime);
                if (p.getMode() == FileChooserParams.MODE_OPEN_MULTIPLE) i.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true);
                try { startActivityForResult(i, REQ_CHOOSER); } catch (Exception e) { chooserCb = null; return false; }
                return true;
            }
        });
        bridge.takeIntent(getIntent());
        String hash = getIntent() != null ? getIntent().getStringExtra("hash") : null;
        if (b != null && hash == null && !isDocWindow()) web.restoreState(b); else web.loadUrl(ORIGIN + "index.html" + (isDocWindow() ? "?win=doc" : "") + (hash != null ? "#" + hash : ""));
    }

    protected boolean isDocWindow() { return false; }

    void grantMedia(PermissionRequest r) {
        List<String> ok = new ArrayList<>();
        for (String res : r.getResources()) {
            if (PermissionRequest.RESOURCE_AUDIO_CAPTURE.equals(res) && checkSelfPermission(Manifest.permission.RECORD_AUDIO) == PackageManager.PERMISSION_GRANTED) ok.add(res);
            else if (PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(res) && checkSelfPermission(Manifest.permission.CAMERA) == PackageManager.PERMISSION_GRANTED) ok.add(res);
        }
        if (ok.isEmpty()) r.deny(); else r.grant(ok.toArray(new String[0]));
    }
    @Override public void onRequestPermissionsResult(int req, String[] perms, int[] res) {
        if (req == REQ_MEDIA && pendingMedia != null) { PermissionRequest r = pendingMedia; pendingMedia = null; grantMedia(r); return; }
        super.onRequestPermissionsResult(req, perms, res);
    }

    boolean external(Uri u) {
        if (u != null && "oos.app".equals(u.getHost())) return false;
        try { startActivity(new Intent(Intent.ACTION_VIEW, u)); } catch (Exception ignored) {}
        return true;
    }

    static final Map<String, String> MIME = new HashMap<>();
    static {
        MIME.put("html", "text/html"); MIME.put("js", "application/javascript"); MIME.put("css", "text/css"); MIME.put("json", "application/json");
        MIME.put("png", "image/png"); MIME.put("jpg", "image/jpeg"); MIME.put("svg", "image/svg+xml"); MIME.put("ico", "image/x-icon");
        MIME.put("woff2", "font/woff2"); MIME.put("woff", "font/woff"); MIME.put("ttf", "font/ttf"); MIME.put("wasm", "application/wasm");
        MIME.put("traineddata", "application/octet-stream"); MIME.put("txt", "text/plain");
        MIME.put("otf", "font/otf"); MIME.put("ttc", "font/collection"); MIME.put("mjs", "application/javascript"); MIME.put("mp4", "video/mp4"); MIME.put("webm", "video/webm"); MIME.put("mp3", "audio/mpeg"); MIME.put("zip", "application/zip"); MIME.put("gif", "image/gif"); MIME.put("webp", "image/webp");
    }
    WebResourceResponse serve(Uri u) {
        if (u == null || !"oos.app".equals(u.getHost())) return null;
        String path = u.getPath(); if (path == null || path.equals("/") || path.isEmpty()) path = "/index.html";
        try {
            InputStream in = getAssets().open("www" + path);
            String ext = path.substring(path.lastIndexOf('.') + 1).toLowerCase();
            String mime = MIME.containsKey(ext) ? MIME.get(ext) : "application/octet-stream";
            String enc = mime.startsWith("text/") || mime.contains("javascript") || mime.contains("json") || mime.contains("svg") ? "utf-8" : null;
            WebResourceResponse r = new WebResourceResponse(mime, enc, in);
            Map<String, String> h = new HashMap<>(); h.put("Access-Control-Allow-Origin", "*"); h.put("Cache-Control", "no-cache");
            r.setResponseHeaders(h);
            return r;
        } catch (Exception e) {
            return new WebResourceResponse("text/plain", "utf-8", 404, "Not Found", null, new java.io.ByteArrayInputStream(new byte[0]));
        }
    }

    @Override protected void onNewIntent(Intent i) {
        super.onNewIntent(i);
        String hash = i.getStringExtra("hash");
        if (hash != null) { web.loadUrl(ORIGIN + "index.html?t=" + System.currentTimeMillis() + "#" + hash); return; }
        if (bridge.takeIntent(i)) web.evaluateJavascript("window.OOS_Host && OOS_Host.pending().then(function(f){f.forEach(function(x){window.OOS.openFile(x)})})", null);
    }
    @Override protected void onActivityResult(int req, int res, Intent data) {
        if (req == REQ_CHOOSER) {
            if (chooserCb == null) return;
            Uri[] out = null;
            if (res == RESULT_OK && data != null) {
                if (data.getClipData() != null) { int n = data.getClipData().getItemCount(); out = new Uri[n]; for (int k = 0; k < n; k++) out[k] = data.getClipData().getItemAt(k).getUri(); }
                else if (data.getData() != null) out = new Uri[]{data.getData()};
            }
            chooserCb.onReceiveValue(out); chooserCb = null; return;
        }
        if (req == REQ_AIWEB) { if (res == RESULT_OK && data != null && data.getStringExtra("answer") != null) web.evaluateJavascript("window.OOS_onAiAnswer && OOS_onAiAnswer(" + org.json.JSONObject.quote(data.getStringExtra("answer")) + ")", null); return; }
        if (req == REQ_CREATE) { bridge.onCreateResult(res == RESULT_OK && data != null ? data.getData() : null); return; }
        super.onActivityResult(req, res, data);
    }
    @Override public void onBackPressed() {
        web.evaluateJavascript("(window.OOS_onBack && window.OOS_onBack()) ? '1' : '0'", v -> { if (!"\"1\"".equals(v)) { if (isDocWindow()) finishAndRemoveTask(); else moveTaskToBack(true); } });
    }
    @Override public void onConfigurationChanged(android.content.res.Configuration newConfig) {
        super.onConfigurationChanged(newConfig);
        applySafeInsets();
        if (web != null && bridge != null) {
            boolean isDex = bridge.isSamsungDeX();
            web.post(() -> web.evaluateJavascript("window.OOS_onDisplayChanged && window.OOS_onDisplayChanged(" + isDex + ")", null));
        }
    }
    @Override protected void onSaveInstanceState(Bundle b) { super.onSaveInstanceState(b); web.saveState(b); }
    @Override protected void onPause() { super.onPause(); web.evaluateJavascript("window.OOS && OOS.flushAutosave && OOS.flushAutosave()", null); }
    @Override protected void onDestroy() {
        if (bridge != null) bridge.stopWebLink();
        if (web != null) { web.destroy(); }
        super.onDestroy();
    }

    void callback(final String id, final String json) {
        if (id == null) return;
        final String js = "window.__oosCb && __oosCb(" + org.json.JSONObject.quote(id) + "," + org.json.JSONObject.quote(json == null ? "null" : json) + ")";
        runOnUiThread(() -> web.evaluateJavascript(js, null));
    }
}
