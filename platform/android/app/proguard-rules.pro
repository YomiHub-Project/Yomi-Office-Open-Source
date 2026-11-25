-keepclassmembers class com.officeopensource.office_open_source.Bridge { @android.webkit.JavascriptInterface <methods>; }
-keepattributes JavascriptInterface

# v0.1.6: LibreOfficeKit JNI (native code looks these up by name)
-keep class org.libreoffice.kit.** { *; }
-keep class com.officeopensource.office_open_source.NativeDocActivity { *; }
