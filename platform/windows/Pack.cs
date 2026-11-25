// Builds the Setup payload: gzip( repeat { int32 pathLen, utf8 path, int64 size, bytes } ) from a folder
using System; using System.IO; using System.IO.Compression; using System.Text;
static class Pack {
    static int Main(string[] a) {
        string src = Path.GetFullPath(a[0]).TrimEnd('\\') + "\\"; int n = 0;
        using (FileStream fs = File.Create(a[1])) using (GZipStream gz = new GZipStream(fs, CompressionMode.Compress)) using (BinaryWriter bw = new BinaryWriter(gz)) {
            foreach (string f in Directory.GetFiles(src, "*", SearchOption.AllDirectories)) {
                byte[] p = Encoding.UTF8.GetBytes(f.Substring(src.Length)); byte[] d = File.ReadAllBytes(f);
                bw.Write(p.Length); bw.Write(p); bw.Write((long)d.Length); bw.Write(d); n++;
            }
        }
        Console.WriteLine("packed " + n + " files -> " + new FileInfo(a[1]).Length + " bytes"); return 0;
    }
}
