library JpegSave;



uses
//    windows,
    graphics,
    jpeg;

type
  vtJPEGInfo=packed record
      jpQuality:longint;
      Progressive:wordbool;
    end;


var jpg:TJpegImage;
    bmp:TBitmap;
//    mf : TMetaFile;
//    mfCnv : TMetaFileCanvas;

//Эта процедура сохраняет в Jpg
  procedure SaveToJpeg (hBitmap:LongWord;
                        FileName:PChar;
                        var Info: vtJPEGInfo);safecall;
  var
    bmHandle:LongWord;
   begin
    jpg := tjpegimage.Create  ;   // Создаём Джепег
    bmp := tbitmap.Create ;       // Создаём битмап

    bmhandle:=bmp.handle;
    bmp.Handle := hbitmap;

    jpg.CompressionQuality := info.jpQuality ; // Устанавливаем качество
    jpg.ProgressiveEncoding := info.Progressive;
    jpg.Assign (bmp); // Ассигнуем
    jpg.SaveToFile (filename); // Сохраняем

    //*** Освободим память
    jpg.Free;
    bmp.Handle := bmhandle;
    bmp.Free;
   end;


exports SaveToJpeg ;

begin
end.
