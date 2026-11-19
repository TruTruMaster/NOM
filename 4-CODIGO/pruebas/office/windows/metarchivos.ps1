# Dibuja cada EMF o WMF de una carpeta con el GDI de Windows (PlayEnhMetaFile: lo mismo con que Word, Excel y
# PowerPoint muestran un metarchivo; no GDI+, que los reinterpreta) a un PNG del tamaño pedido, sobre blanco.
# Un WMF pasa a EMF con SetWinMetaFileBits y su caja del encabezado «placeable», como lo hace Office.
# Lo usa pruebas/office/metarchivos_real.py.
#   powershell -File metarchivos.ps1 <carpeta con .emf/.wmf> <carpeta de salida> <lista.txt: nombre<TAB>ancho<TAB>alto>
param([string]$Entrada, [string]$Salida, [string]$Lista)
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
public static class Gdi {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int l, t, r, b; }
  [StructLayout(LayoutKind.Sequential)] public struct METAFILEPICT { public int mm, xExt, yExt; public IntPtr hMF; }
  [StructLayout(LayoutKind.Sequential)] public struct BITMAPINFOHEADER { public int biSize, biWidth, biHeight; public short biPlanes, biBitCount; public int biCompression, biSizeImage, biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant; }
  [DllImport("gdi32.dll")] static extern IntPtr GetEnhMetaFileW([MarshalAs(UnmanagedType.LPWStr)] string f);
  [DllImport("gdi32.dll")] static extern IntPtr SetWinMetaFileBits(uint n, byte[] b, IntPtr hdc, ref METAFILEPICT mfp);
  [DllImport("gdi32.dll")] static extern bool PlayEnhMetaFile(IntPtr hdc, IntPtr hemf, ref RECT r);
  [DllImport("gdi32.dll")] static extern bool DeleteEnhMetaFile(IntPtr h);
  [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr hdc);
  [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr hdc);
  [DllImport("gdi32.dll")] static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFOHEADER bmi, uint usage, out IntPtr bits, IntPtr sec, uint off);
  [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr hdc, IntPtr o);
  [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr o);
  [DllImport("gdi32.dll")] static extern bool PatBlt(IntPtr hdc, int x, int y, int w, int h, uint rop);
  public static void Dibujar(string ruta, string salida, int w, int h) {
    IntPtr hemf;
    if (ruta.ToLower().EndsWith(".wmf")) {
      byte[] b = File.ReadAllBytes(ruta);
      int off = 0; METAFILEPICT p = new METAFILEPICT(); p.mm = 8;
      if (b.Length > 22 && BitConverter.ToUInt32(b, 0) == 0x9AC6CDD7) {
        short l = BitConverter.ToInt16(b, 6), t = BitConverter.ToInt16(b, 8), r = BitConverter.ToInt16(b, 10), bo = BitConverter.ToInt16(b, 12);
        int k = BitConverter.ToUInt16(b, 14); if (k == 0) k = 1440;
        p.xExt = (int)Math.Round(Math.Abs(r - l) * 2540.0 / k); p.yExt = (int)Math.Round(Math.Abs(bo - t) * 2540.0 / k);
        off = 22;
      }
      byte[] c = new byte[b.Length - off]; Array.Copy(b, off, c, 0, c.Length);
      hemf = SetWinMetaFileBits((uint)c.Length, c, IntPtr.Zero, ref p);
    } else hemf = GetEnhMetaFileW(ruta);
    if (hemf == IntPtr.Zero) throw new Exception("GDI no lo abre");
    IntPtr dc = CreateCompatibleDC(IntPtr.Zero);
    BITMAPINFOHEADER bi = new BITMAPINFOHEADER(); bi.biSize = 40; bi.biWidth = w; bi.biHeight = -h; bi.biPlanes = 1; bi.biBitCount = 32;
    IntPtr bits; IntPtr dib = CreateDIBSection(dc, ref bi, 0, out bits, IntPtr.Zero, 0);
    IntPtr viejo = SelectObject(dc, dib);
    PatBlt(dc, 0, 0, w, h, 0x00FF0062);
    RECT rc = new RECT(); rc.l = 0; rc.t = 0; rc.r = w; rc.b = h;
    PlayEnhMetaFile(dc, hemf, ref rc);
    using (Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppRgb)) {
      BitmapData d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
      byte[] tmp = new byte[w * h * 4]; Marshal.Copy(bits, tmp, 0, tmp.Length);
      for (int y = 0; y < h; y++) Marshal.Copy(tmp, y * w * 4, d.Scan0 + y * d.Stride, w * 4);
      bmp.UnlockBits(d);
      bmp.Save(salida, ImageFormat.Png);
    }
    SelectObject(dc, viejo); DeleteObject(dib); DeleteDC(dc); DeleteEnhMetaFile(hemf);
  }
}
"@
New-Item -ItemType Directory -Force -Path $Salida | Out-Null
$ok = 0; $mal = 0
foreach ($linea in Get-Content $Lista) {
  $p = $linea -split "`t"
  if ($p.Count -lt 3) { continue }
  try { [Gdi]::Dibujar((Resolve-Path (Join-Path $Entrada $p[0])).Path, (Join-Path (Resolve-Path $Salida).Path ($p[0] + ".png")), [int]$p[1], [int]$p[2]); $ok++ }
  catch { Write-Output ("MAL " + $p[0] + " " + $_.Exception.Message); $mal++ }
}
Write-Output ("listos " + $ok + " fallaron " + $mal)
