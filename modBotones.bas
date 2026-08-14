Attribute VB_Name = "modBotones"
Public Const BLACKNESS = &H42
' public Const CAPTUREBLT = ???
Public Const DSTINVERT = &H550009
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
' public Const NOMIRRORBITMAP = ???
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const PATCOPY = &HF00021
Public Const PATINVERT = &H5A0049
Public Const PATPAINT = &HFB0A09
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const WHITENESS = &HFF0062

#If Win32 Then
        Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X _
        As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As _
        Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As _
        Long, ByVal dwRop As Long) As Long
        #Else
        Declare Function BitBlt Lib "GDI" (ByVal hDestDC%, ByVal X%, ByVal _
        Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, _
        ByVal YSrc%, ByVal dwRop&) As Integer
        #End If

Public Sub ActivarBoton(forma As Form, Estado As String, formahDC As Long, destino As PictureBox)
Dim tipo As String, i As Integer
tipo = destino.Tag
If InStr(tipo, "/") Then
    i = InStr(tipo, "/L")
    lBot = Mid(tipo, i + 2)
    tipo = Mid(tipo, 1, i - 1)
    i = InStr(tipo, "/T")
    tBot = Mid(tipo, i + 2)
    tipo = Mid(destino.Tag, 1, i - 1)
End If
With forma
If tipo = "miniok" Then
    .picAll.Picture = .BotonesImages.ListImages(1).Picture
ElseIf tipo = "minino" Then
    .picAll.Picture = .BotonesImages.ListImages(2).Picture
ElseIf tipo = "ok" Then
    .picAll.Picture = .BotonesImages.ListImages(3).Picture
ElseIf tipo = "no" Then
    .picAll.Picture = .BotonesImages.ListImages(4).Picture
ElseIf tipo = "agregar" Then
    .picAll.Picture = .BotonesImages.ListImages(5).Picture
ElseIf tipo = "facturar" Then
    .picAll.Picture = .BotonesImages.ListImages(6).Picture
ElseIf tipo = "borrartodo" Then
    .picAll.Picture = .BotonesImages.ListImages(7).Picture
ElseIf tipo = "miniclientes" Then
    .picAll.Picture = .BotonesImages.ListImages(8).Picture
ElseIf tipo = "miniinventario" Then
    .picAll.Picture = .BotonesImages.ListImages(9).Picture
ElseIf tipo = "minireportes" Then
    .picAll.Picture = .BotonesImages.ListImages(10).Picture
ElseIf tipo = "minidolar" Then
    .picAll.Picture = .BotonesImages.ListImages(11).Picture
ElseIf tipo = "miniconfig" Then
    .picAll.Picture = .BotonesImages.ListImages(12).Picture
ElseIf tipo = "chart" Then
    .picAll.Picture = .BotonesImages.ListImages(13).Picture
ElseIf tipo = "printer" Then
    .picAll.Picture = .BotonesImages.ListImages(14).Picture
ElseIf tipo = "miniprinter" Then
    .picAll.Picture = .BotonesImages.ListImages(15).Picture
ElseIf tipo = "export" Then
    .picAll.Picture = .BotonesImages.ListImages(16).Picture
ElseIf tipo = "miniadd" Then
    .picAll.Picture = .BotonesImages.ListImages(17).Picture
ElseIf tipo = "minicalc" Then
    .picAll.Picture = .BotonesImages.ListImages(18).Picture
ElseIf tipo = "clienteadd" Then
    .picAll.Picture = .BotonesImages.ListImages(19).Picture
ElseIf tipo = "clientedel" Then
    .picAll.Picture = .BotonesImages.ListImages(20).Picture
ElseIf tipo = "clienteedit" Then
    .picAll.Picture = .BotonesImages.ListImages(21).Picture
ElseIf tipo = "invadd" Then
    .picAll.Picture = .BotonesImages.ListImages(22).Picture
ElseIf tipo = "invdel" Then
    .picAll.Picture = .BotonesImages.ListImages(23).Picture
ElseIf tipo = "invedit" Then
    .picAll.Picture = .BotonesImages.ListImages(24).Picture
ElseIf tipo = "dbpath" Then
    .picAll.Picture = .BotonesImages.ListImages(25).Picture
ElseIf tipo = "autorizar" Then
    .picAll.Picture = .BotonesImages.ListImages(26).Picture
ElseIf tipo = "cierrez" Then
    .picAll.Picture = .BotonesImages.ListImages(27).Picture
ElseIf tipo = "cierrex" Then
    .picAll.Picture = .BotonesImages.ListImages(28).Picture
ElseIf tipo = "find" Then
    .picAll.Picture = .BotonesImages.ListImages(29).Picture
ElseIf tipo = "microfind" Then
    .picAll.Picture = .BotonesImages.ListImages(30).Picture
ElseIf tipo = "clientes" Then
    .picAll.Picture = .BotonesImages.ListImages(31).Picture
ElseIf tipo = "inventario" Then
    .picAll.Picture = .BotonesImages.ListImages(32).Picture
End If

wBot = (.picAll.Width / 2): hBot = (.picAll.Height / 2)
destino.Width = wBot
destino.Height = hBot
.mask.Width = wBot
.mask.Height = hBot
.mask_invert.Width = wBot
.mask_invert.Height = hBot
.picUp.Width = wBot
.picUp.Height = hBot
.picDo.Width = wBot
.picDo.Height = hBot
.mask.PaintPicture .picAll.Picture, 0, 0, .picAll.Width / 2, .picAll.Height / 2, .picAll.Width / 2, .picAll.Height / 2, .picAll.Width / 2, .picAll.Height / 2
.mask_invert.PaintPicture .picAll.Picture, 0, 0, .picAll.Width / 2, .picAll.Height / 2, 0, .picAll.Height / 2, .picAll.Width / 2, .picAll.Height / 2
.picUp.PaintPicture .picAll.Picture, 0, 0, .picAll.Width / 2, .picAll.Height / 2, 0, 0, .picAll.Width / 2, .picAll.Height / 2
.picDo.PaintPicture .picAll.Picture, 0, 0, .picAll.Width / 2, .picAll.Height / 2, .picAll.Width / 2, 0, .picAll.Width / 2, .picAll.Height / 2

If Estado = "do" Then
    destino.Left = lBot + 1
    destino.Top = tBot + 1
    destino.AutoRedraw = False
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, formahDC, lBot + 1, tBot + 2, SRCCOPY) 'copiar del fondo
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .mask_invert.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(.picDo.hDC, 0, 0, wBot, hBot, .mask.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .picDo.hDC, 0, 0, SRCPAINT)
    destino.AutoRedraw = True
ElseIf Estado = "up" Then
    destino.Left = lBot - 1
    destino.Top = tBot - 1
    destino.AutoRedraw = False
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, formahDC, lBot - 1, tBot - 2, SRCCOPY) 'copiar del fondo
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .mask_invert.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(.picUp.hDC, 0, 0, wBot, hBot, .mask.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .picUp.hDC, 0, 0, SRCPAINT)
    destino.AutoRedraw = True
ElseIf Estado = "ini" Then
    lBot = destino.Left
    tBot = destino.Top
    destino.Tag = destino.Tag & "/T" & destino.Top & "/L" & destino.Left
    destino.AutoRedraw = True
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, formahDC, lBot, tBot, SRCCOPY) 'copiar del fondo
    Suc% = BitBlt(.picUp.hDC, 0, 0, wBot, hBot, .mask.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .mask_invert.hDC, 0, 0, SRCAND)
    Suc% = BitBlt(destino.hDC, 0, 0, wBot, hBot, .picUp.hDC, 0, 0, SRCPAINT)
End If
End With
End Sub
