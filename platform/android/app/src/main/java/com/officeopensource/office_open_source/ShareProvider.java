package com.officeopensource.office_open_source;

import android.content.ContentProvider;
import android.content.ContentValues;
import android.database.Cursor;
import android.database.MatrixCursor;
import android.net.Uri;
import android.os.ParcelFileDescriptor;
import android.provider.OpenableColumns;
import android.webkit.MimeTypeMap;

import java.io.File;
import java.io.FileNotFoundException;

/** Minimal read-only provider for sharing exported files from cache/share (no AndroidX FileProvider needed). */
public class ShareProvider extends ContentProvider {
    static final String AUTH = "com.officeopensource.office_open_source.share";
    File file(Uri u) { return new File(new File(getContext().getCacheDir(), "share"), new File(u.getLastPathSegment()).getName()); }
    @Override public boolean onCreate() { return true; }
    @Override public ParcelFileDescriptor openFile(Uri u, String mode) throws FileNotFoundException { return ParcelFileDescriptor.open(file(u), ParcelFileDescriptor.MODE_READ_ONLY); }
    @Override public String getType(Uri u) { String n = u.getLastPathSegment(); String e = n.substring(n.lastIndexOf('.') + 1).toLowerCase(); String m = MimeTypeMap.getSingleton().getMimeTypeFromExtension(e); return m == null ? "application/octet-stream" : m; }
    @Override public Cursor query(Uri u, String[] p, String s, String[] a, String o) { File f = file(u); MatrixCursor c = new MatrixCursor(new String[]{OpenableColumns.DISPLAY_NAME, OpenableColumns.SIZE}); c.addRow(new Object[]{f.getName(), f.length()}); return c; }
    @Override public Uri insert(Uri u, ContentValues v) { return null; }
    @Override public int delete(Uri u, String s, String[] a) { return 0; }
    @Override public int update(Uri u, ContentValues v, String s, String[] a) { return 0; }
}
