PARAM (
    
    [string] $ZipFilesPath = "C:\Archivos recibidos",
    [string] $ZipFilesDest = "C:\Archivos descomprimidos"

)

$Shell = New-Object -com Shell.Application
$Location = $Shell.NameSpace($ZipFilesDest)

$ZipFiles = Get-ChildItem $ZipFilesPath -Recurse -Include *.ZIP

# Aquí iniciamos el Bucle para reccorrer los archivos ZIP y descomprimirlos en la carpeta Dest (Destino)

$i

foreach ($ZipFile in $ZipFiles) {
    
    Write-Progress -Activity "Descromprimiendo en $($ZipFilesDest)" -PercentComplete (($i / ($ZipFiles.count +1))*100) -CurrentOperation $ZipFile.FullName -Status " Archivo $($ZipFiles.count)"

    $ZipFolder = $Shell.NameSpace($ZipFile.fullname)

    $Location.Copyhere($ZipFolder.items(), 1040)
    $i++
}