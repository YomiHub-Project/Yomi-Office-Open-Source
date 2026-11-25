package com.officeopensource.office_open_source;

/** v0.0.6: every document opened from the app (or from another app while a document is open) gets its own window/task. */
public class DocActivity extends MainActivity {
    @Override protected boolean isDocWindow() { return true; }
}
