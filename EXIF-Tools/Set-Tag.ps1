[Reflection.Assembly]::LoadFrom( ("C:\Users\Luca\Documents\MPTag\taglib-sharp.dll") )
$media = [TagLib.File]::Create("C:\temp\photo.jpg")
$media.Tag.Title = "Test Title"
$media.save()
