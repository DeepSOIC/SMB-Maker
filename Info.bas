Attribute VB_Name = "Info"
'SMB Maker Info module
'No code, comments only
'
'Resource dll contents:

'102=mpng.dll



'------------------------BMP FILE DESCRIPTION-----------------------------------------------
'Bitmap file header      14 bait;
'Bitmap info header       40 bait;
'Palitra         ot kol tsvetov (n*4);
'Baitovyi massiv     ot razmerov i kolichestva bit na piksel;
'
'
'Bitmap file header
'Sleduyuschie polya:
'WORD bfType             hranit simvoly "BM". Eto kod formata
'DWORD bfSizeJ               Obschii razmer faila v baitah
'WORD bfReserved1 = 0
'WORD bfReserved2 = 0
'DWORD bfOffBits             adres bitovogo massiva v dannom faile
'
'Bitmap info header
'
'DWORD biSize                Razmer zagolovka, = 40
'LONG biWidth                Shirina rastra (matritsy) v pikselah
'LONG biHeight               Vysota
'WORD biPlanes = 1
'WORD biBitCount             bit na piksel
'DWORD biCompression = 0
'DWORD biSizeImage           razmer v baitah bitovogo massiva rastra
'LONG biXPelsPerMeter            razreshenie po X v pikselah na metr
'LONG biYPelsPerMeter            - po Y
'DWORD biClrUsed             esli = 0, to ispol'zuetsya maksimal'noe kolichestvo tsvetov
'DWORD biClrImprtant         = 0, esli biClrUsed = 0
'
'Zatem pomeschaetsya palitra. Kazhdaya zapis' soderzhit 4 polya:
'BYTE Blue
'BYTE Green
'BYTE Red
'BYTE rgbReserved
'Palitra otsutstvuet, esli chislo bit na piksel = 24. Takzhe palitra ne nuzhna dlya nekotoryh tsvetovyh formatov 16 i 32 bit na piksel.
'
'Dalee zapisyvaetsya rastr v vide bitovogo (tochnee baitovogo) massiva. Posledovatel'no zapisyvayutsya baity strok rastra. Kolichestvo bait v stroke dolzhno byt' kratnym 4. (vyravnivanie na granitsu dvoinogo slova)
'drugie formaty mozhno naiti v "Born G. Formaty dannyh. - SPb.: BHV, 1995. - 472 s."
'




'----------------------------


