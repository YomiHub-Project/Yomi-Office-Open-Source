package com.officeopensource.office_open_source;

import android.content.Context;
import android.os.Environment;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.Inet4Address;
import java.net.InetAddress;
import java.net.NetworkInterface;
import java.net.ServerSocket;
import java.net.Socket;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.UUID;

/** v0.1.6 - AirDrop / QuickShare over Wi-Fi: real two-way LAN file transfer (same protocol as the Windows build). */
final class LanShare {
    static final int PORT = 8765;
    static ServerSocket server;
    static final HashMap<String, Object[]> files = new HashMap<>();
    static final List<String> received = new ArrayList<>();
    static Context ctx;

    static File inbox() {
        // Android 11+: apps may create files in the public Download folder; older versions use the app's own external folder
        File d = android.os.Build.VERSION.SDK_INT >= 30 ? new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "Office Open Source") : new File(ctx.getExternalFilesDir(null), "Nhan");
        if (!d.mkdirs() && !d.isDirectory()) d = new File(ctx.getExternalFilesDir(null), "Nhan");
        d.mkdirs(); return d;
    }

    static List<String> ips() {
        List<String> r = new ArrayList<>();
        try {
            for (NetworkInterface ni : Collections.list(NetworkInterface.getNetworkInterfaces())) {
                if (!ni.isUp() || ni.isLoopback()) continue;
                for (InetAddress a : Collections.list(ni.getInetAddresses())) if (a instanceof Inet4Address && !a.isLoopbackAddress()) r.add(a.getHostAddress());
            }
        } catch (Exception ignored) {}
        return r;
    }

    static synchronized boolean start(Context c) {
        ctx = c.getApplicationContext();
        if (server != null) return true;
        try {
            server = new ServerSocket(PORT);
            new Thread(() -> { while (true) { try { final Socket s = server.accept(); new Thread(() -> serve(s)).start(); } catch (Exception e) { return; } } }, "lanshare").start();
            return true;
        } catch (Exception e) { server = null; return false; }
    }

    static String share(Context c, String name, byte[] data) throws Exception {
        boolean ok = start(c);
        String tok = UUID.randomUUID().toString().replace("-", "").substring(0, 10);
        synchronized (files) { files.put(tok, new Object[]{ name, data }); }
        JSONArray urls = new JSONArray(), recv = new JSONArray();
        for (String ip : ips()) { urls.put("http://" + ip + ":" + PORT + "/s/" + tok); recv.put("http://" + ip + ":" + PORT + "/"); }
        JSONObject o = new JSONObject(); o.put("ok", ok); o.put("token", tok); o.put("urls", urls); o.put("receive", recv); o.put("inbox", inbox().getAbsolutePath());
        return o.toString();
    }

    static String esc(String s) { return s == null ? "" : s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;"); }
    static String page(String title, String body) {
        return "<!doctype html><html lang=vi><head><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'><title>" + esc(title) + "</title>"
            + "<style>body{font-family:system-ui,Roboto,sans-serif;background:#f3f6fb;margin:0;display:flex;min-height:100vh;align-items:center;justify-content:center}.c{background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.12);padding:28px;max-width:420px;width:90%;text-align:center}a.b,button{display:inline-block;background:#185abd;color:#fff;border:0;border-radius:10px;padding:12px 22px;font-size:16px;text-decoration:none;margin-top:14px}small{color:#64748b}</style></head><body><div class=c>"
            + body + "</div></body></html>";
    }

    static void serve(Socket s) {
        try (Socket sock = s; InputStream in = sock.getInputStream(); OutputStream out = sock.getOutputStream()) {
            sock.setSoTimeout(30000);
            ByteArrayOutputStream hb = new ByteArrayOutputStream(); int last4 = 0, b;
            while ((b = in.read()) >= 0) { hb.write(b); last4 = (last4 << 8) | b; if (last4 == 0x0D0A0D0A || hb.size() > 65536) break; }
            String head = hb.toString("UTF-8"); String[] lines = head.split("\r\n"); String[] rq = lines[0].split(" "); if (rq.length < 2) return;
            String method = rq[0], url = rq[1], path = url.split("\\?")[0]; long clen = 0;
            for (String l : lines) if (l.toLowerCase().startsWith("content-length:")) clen = Long.parseLong(l.substring(15).trim());
            if ("GET".equals(method) && path.startsWith("/s/")) {
                String[] p = path.substring(3).split("/"); Object[] f; synchronized (files) { f = files.get(p[0]); }
                if (f == null) { write(out, 404, "text/html; charset=utf-8", page("404", "<h2>Liên kết đã hết hạn</h2>").getBytes("UTF-8"), null); return; }
                String name = (String) f[0]; byte[] data = (byte[]) f[1];
                if (p.length > 1 && "dl".equals(p[1])) { write(out, 200, "application/octet-stream", data, "attachment; filename*=UTF-8''" + URLEncoder.encode(name, "UTF-8").replace("+", "%20")); return; }
                write(out, 200, "text/html; charset=utf-8", page(name, "<div style='font-size:48px'>📄</div><h2>" + esc(name) + "</h2><small>" + (data.length / 1024) + " KB • gửi từ Office Open Source (" + esc(android.os.Build.MODEL) + ")</small><br><a class=b href='/s/" + p[0] + "/dl'>Tải xuống</a>").getBytes("UTF-8"), null);
                return;
            }
            if ("GET".equals(method) && ("/".equals(path) || "/index.html".equals(path))) {
                String js = "<script>function up(){var f=document.getElementById('f').files;if(!f.length)return;var n=0,m=document.getElementById('m');m.textContent='Đang gửi…';[].forEach.call(f,function(x){var r=new XMLHttpRequest();r.open('POST','/up?name='+encodeURIComponent(x.name));r.onload=function(){n++;m.textContent='Đã gửi '+n+'/'+f.length+' tệp';};r.send(x);});}</script>";
                write(out, 200, "text/html; charset=utf-8", page("Gửi tệp", "<h2>Gửi tệp tới " + esc(android.os.Build.MODEL) + "</h2><small>Tệp được lưu vào Download/Office Open Source</small><br><input id=f type=file multiple style='margin-top:16px'><br><button onclick=up()>Gửi</button><p id=m></p>" + js).getBytes("UTF-8"), null);
                return;
            }
            if ("POST".equals(method) && "/up".equals(path)) {
                String name = "tep-nhan-" + System.currentTimeMillis();
                int qi = url.indexOf('?'); if (qi > 0) for (String kv : url.substring(qi + 1).split("&")) if (kv.startsWith("name=")) name = URLDecoder.decode(kv.substring(5), "UTF-8");
                name = new File(name).getName().replaceAll("[\\\\/:*?\"<>|]", "_");
                if (clen <= 0 || clen > 1024L * 1024 * 1024) { write(out, 413, "text/plain", "too large".getBytes(), null); return; }
                File dest = new File(inbox(), name); int k = 1;
                while (dest.exists()) { String n = name; int dot = n.lastIndexOf('.'); dest = new File(inbox(), (dot > 0 ? n.substring(0, dot) : n) + " (" + (k++) + ")" + (dot > 0 ? n.substring(dot) : "")); }
                try (FileOutputStream fo = new FileOutputStream(dest)) { byte[] buf = new byte[65536]; long left = clen; while (left > 0) { int n = in.read(buf, 0, (int) Math.min(buf.length, left)); if (n <= 0) break; fo.write(buf, 0, n); left -= n; } }
                synchronized (received) { received.add(dest.getAbsolutePath()); }
                write(out, 200, "application/json", "{\"ok\":true}".getBytes(), null);
                return;
            }
            write(out, 404, "text/plain", "not found".getBytes(), null);
        } catch (Exception ignored) {}
    }

    static void write(OutputStream o, int code, String type, byte[] body, String disp) throws Exception {
        String h = "HTTP/1.1 " + code + " OK\r\nContent-Type: " + type + "\r\nContent-Length: " + body.length + "\r\n" + (disp != null ? "Content-Disposition: " + disp + "\r\n" : "") + "Connection: close\r\nCache-Control: no-store\r\n\r\n";
        o.write(h.getBytes("UTF-8")); o.write(body); o.flush();
    }
}
