#Requires -Version 7

[CmdletBinding()]
param 
(
    [switch]$GLS,
    [switch]$Textual,
    [switch]$TSV,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("EN","FR","ES","DE","RU","IT")]
    [string]$Dil,
    [switch]$Hunspell
)

function ENTR {
    param (
        [xml]$parca
    )
    Add-Type -AssemblyName $PSScriptRoot\Modul\RtfPipe.dll

    $cikti = foreach ($i in $parca.FONO.KELIME) {
        $baslik = $i.SOZCUK
        if ($baslik.Length -lt 1) { continue }
        $tanim = [System.String]$i.ACIKLAMA
        if ($tanim.Length -lt 1) { continue }
        $tanim_html = [RtfPipe.Rtf]::ToHtml($tanim)
        $tanim_html = $tanim_html -replace
                        "<p[^>]+>", "<p>" -replace
                        "<strong[^>]+>", "<strong>" -replace
                        "<span[^>]+>", "<span>" -replace
                        "<em[^>]+>", "<em>" -replace
                        "<div[^>]+>", "<div>"
        @{
            BASLIK = $baslik.ReplaceLineEndings("").Trim()
            TANIM  = $tanim_html.ReplaceLineEndings("").Trim()
        }
    }
    return $cikti
}

function findIndex
{
    param(
        [System.Byte[]]$array,
        [int]$start
    )
    $sep = @(0x0D, 0x00, 0x0A, 0x00)
    $MAX = 128 + $start
    for ($i = $start; $i -lt $MAX; $i++)
    {
        if ($array[$i] -eq $sep[0])
        {
            if (
                $array[$i+1] -eq $sep[1] -and
                $array[$i+2] -eq $sep[2] -and
                $array[$i+3] -eq $sep[3]
            )
            {
                return $i
            }
        }
    }
    return -1
}
function EuroDictXP
{
    param(
        [string]$fileName,
        [bool]$add_abbrv
    )
    Add-Type -AssemblyName $PSScriptRoot\Modul\RtfPipe.dll
    
    $offset = 0x4D5
    [System.Byte[]]$byteArray = [System.IO.File]::ReadAllBytes($fileName)
    [System.Byte[]]$byteArray = $byteArray[$offset..$byteArray.Length]
    [System.Collections.ArrayList]$arr     = @()
    [System.Collections.ArrayList]$hwlist  = @()
    [System.Collections.ArrayList]$deflist = @()
    [System.Collections.ArrayList]$offsets = @()
    $enc8  = [Text.Encoding]::UTF8
    $enc16 = [Text.Encoding]::Unicode

    $lastIndex = 0
    while ($true) 
    {
        $skip = 0
        $idx = findIndex -array $byteArray -start $lastIndex
        if ($idx -eq -1)
        {
            break
        }

        if($byteArray[$lastIndex] -eq 0x00)
        {
            $skip = 8
        }
        $word = $enc16.GetString($byteArray[($lastIndex+$skip)..($idx-1)])
        [void]$hwlist.Add($word)
        $lastIndex = $idx + 4
    }
    Write-Host "*** $(Split-Path $fileName -Leaf) için tüm madde başlıkları bulundu."

    for(($i = $lastIndex) ; ($i -lt $byteArray.Length) ; ($i++))
    {
        if($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0xDA) # tek başına 32869,32763
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x9C)
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x01) # ikisini ekleyince 33129,32778
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x5E) # bunu da ekleyince 33408,32778
        {
            [void]$offsets.Add($i)
        }
    }
    for(($i = 0) ; ($i -lt $offsets.Count) ; ($i++))
    {
        try
        {
            if($i -eq $offsets.Count-1)
            {
                $inputStream   = [System.IO.MemoryStream]::new($byteArray[$offsets[$i]..($byteArray.Length)])
            }
            elseif($i -eq $offsets.Count-2)
            {
                # sondan bir önceki için 2 yerine 1 atla
                $inputStream   = [System.IO.MemoryStream]::new($byteArray[$offsets[$i]..($offsets[$i+1])])
            }
            else
            {
                # yanlış offset'lerden kaçınmak için 2 sonraki offseti bitiş kısmı olarak geç
                # bir miktar yavaşlık getirebilir ancak sürekli dizi sonunu parça sonu olarak göstermekten kat kat hızlı
                $inputStream   = [System.IO.MemoryStream]::new($byteArray[$offsets[$i]..($offsets[$i+2])])
            }
            # $inputStream   = [System.IO.MemoryStream]::new($byteArray[$offsets[$i]..($byteArray.Length)])
            $outputStream      = [System.IO.MemoryStream]::new()
            $compressionStream = [System.IO.Compression.ZLibStream]::new($inputStream, [System.IO.Compression.CompressionMode]::Decompress)
            $compressionStream.CopyTo($outputStream)
            $out = $outputStream.ToArray()
            
            $decoded = $enc8.GetString($out[4..($out.Length-3)])
            if ($decoded.StartsWith("{"))
            {
                $html    = [RtfPipe.Rtf]::ToHtml([System.String]$decoded)
                $html = $html -replace 
                        "<p[^>]+><br></p>", "" -replace
                        "<div[^>]+>", "<div>" -replace
                        "<p[^>]+>", "<p>" -replace
                        "<strong[^>]+>", "<strong>" -replace
                        "<span[^>]+>", "<span>" -replace
                        "<em[^>]+>", "<em>"
                $chars = foreach ($c in $html.ToCharArray()) {
                    if (-not [char]::IsControl($c))
                    {
                        $c
                    }
                }
                $temiz_html = $chars -join ""
                [void]$deflist.Add($temiz_html)
            }
        }
        catch
        {
            # Write-Host $_
        }
        finally
        {
            $inputStream.Dispose() ; $outputStream.Dispose() ; $compressionStream.Dispose()
        }
    }
    Write-Host "*** $(Split-Path $fileName -Leaf) için tüm tanım gövdeleri bulundu."
    if($add_abbrv)
    {
        [void]$arr.Add(@{ BASLIK = "Kısaltmalar" ; TANIM = $deflist[0].ReplaceLineEndings("").Trim()})
    }
    for ($i = 0; $i -lt $hwlist.Count; $i++)
    {
        [void]$arr.Add(@{ BASLIK = $hwlist[$i].ReplaceLineEndings("").Trim() ; TANIM = $deflist[($i+1)].ReplaceLineEndings("").Trim()})
    }
    return $arr
}

if (-not ($GLS -or $TSV -or $Textual)) {
    Write-Host "[!] En az bir tane çıktı türü belirtmelisiniz."
    exit
}

# https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
switch ($Dil) {
    "EN"
    {
        $TR_ING    = [xml](Get-Content $PSScriptRoot\SozlukDosyalari\TR_ING.xml -Encoding ([System.Text.Encoding]::GetEncoding(1254)))
        $ING_TR    = [xml](Get-Content $PSScriptRoot\SozlukDosyalari\ING_TR.xml -Encoding ([System.Text.Encoding]::GetEncoding(28599)))
        $isim      = "Fono İngilizce⇄Türkçe"
        $dosyaismi = "fono_en_tr"

        $trden_ = ENTR $TR_ING
        Write-Host "*** TR_ING.xml için tüm madde başlıkları/tanım gövdeleri bulundu."
        $trye_  = ENTR $ING_TR
        Write-Host "*** ING_TR.xml için tüm madde başlıkları/tanım gövdeleri bulundu."
        if ($Hunspell) {
            Add-Type -AssemblyName $PSScriptRoot\Modul\HunspellWordForms.dll
            $sozluk_tr = [WordForms]::new("$PSScriptRoot\Hunspell\tr_TR.dic")
            $sozluk_en = [WordForms]::new("$PSScriptRoot\Hunspell\en_US.dic")
            $trden = foreach ($i in $trden_)
            {
                $cekim = ($sozluk_tr.GetWordForms($i.BASLIK, $true, $false, $true)).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
            }
            Write-Host "   *** TR_ING.xml için tüm sözcük sonları bulundu."
            $trye = foreach ($i in $trye_)
            {
                $cekim = ($sozluk_en.GetWordForms($i.BASLIK, $true, $false, $true)).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
            }
            Write-Host "   *** ING_TR.xml için tüm sözcük sonları bulundu."
        }
    }

    "FR"
    {
        $isim      = "Fono Fransızca⇄Türkçe"
        $dosyaismi = "fono_fr_tr"

        $trden_ = EuroDictXP -fileName $PSScriptRoot\SozlukDosyalari\TURFRE_P.KDD -add_abbrv $true
        $trye_  = EuroDictXP -fileName $PSScriptRoot\SozlukDosyalari\FRETUR_P.KDD -add_abbrv $false
        if ($Hunspell) {
            Add-Type -AssemblyName $PSScriptRoot\Modul\HunspellWordForms.dll
            $sozluk_tr = [WordForms]::new("$PSScriptRoot\Hunspell\tr_TR.dic")
            $sozluk_fr = [WordForms]::new("$PSScriptRoot\Hunspell\fr_FR.dic")
            $trden = foreach ($i in $trden_)
            {
                $cekim = ($sozluk_tr.GetWordForms($i.BASLIK, $true, $false, $true)).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
            }
            Write-Host "   *** TURFRE_P.KDD için tüm sözcük sonları bulundu."
            $trye = foreach ($i in $trye_)
            {
                $cekim = ($sozluk_fr.GetWordForms($i.BASLIK, $true, $false, $true)).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
            }
            Write-Host "   *** FRETUR_P.KDD için tüm sözcük sonları bulundu."
        }
    }

    { @("ES","DE","RU","IT") -contains $_ }
    { 
        Write-Host "*** Bu dil test edilemediği için işlenemiyor.`n*** Issues kısmından iletişime geçerek dili eklemek için yardım edebilirsiniz."
        exit
    }
}
if ($Hunspell) {
    $eklenmis = $trden + $trye
}
else {
    $eklenmis = $trden_ + $trye_
}

if ($GLS)
{
    $ustbilgi = @(
    ""
    "#stripmethod=keep"
    "#sametypesequence=h"
    "#bookname=$($isim) Sözlük"
    "#author=Fono Yayınları"
    "#description=$($isim) Sözlük"
    ) -join "`n"
    $ustbilgi += "`n`n"
    Add-Content -Path "$($dosyaismi).gls" -Encoding UTF8 -NoNewline -Value $ustbilgi
    $o = foreach ($i in $eklenmis) {
        if($i.ContainsKey("CEKIM") -and $i.CEKIM.Count -gt 0)
        {
            $_cekim = foreach ($c in $i.CEKIM) {
                if ($c -ine $i.BASLIK)
                {
                    $c
                }
            }
            $b = $i.BASLIK + "|" + ($_cekim -join "|")
            "$($b)`n$($i.TANIM)`n`n"
        }
        else
        {
            "$($i.BASLIK)`n$($i.TANIM)`n`n"
        }
    }   
    # $o = $ustbilgi + $eklenmis -> nedense her maddeden önce bir boşluk ekliyor
    $gls_fname = "$($dosyaismi).gls"
    $o | Out-File -FilePath $gls_fname -Encoding UTF8 -NoNewline -Append
    Write-Host "*** $($gls_fname) yazıldı." -BackgroundColor Green -ForegroundColor Black
}
if ($Textual) {
    [xml]$sdxml = New-Object System.Xml.XmlDocument

    [System.Xml.XmlNode]$header = $sdxml.CreateXmlDeclaration("1.0", "UTF-8", $null)   
    # [System.Xml.XmlAttribute]$namespace = $sdxml.CreateAttribute("xmlns", "xi", "http://www.w3.org/2003/XInclude")
    
    $root = $sdxml.CreateElement("stardict")
    # $root.Attributes.Append($namespace)

    $info     = $sdxml.CreateElement("info")
    $version  = $sdxml.CreateElement("version")     ; $version.InnerText  = "3.0.0"           ; [void]$info.AppendChild($version)
    $bookname = $sdxml.CreateElement("bookname")    ; $bookname.InnerText = "$($isim) Sözlük" ; [void]$info.AppendChild($bookname)
    $author   = $sdxml.CreateElement("author")      ; $author.InnerText   = "Fono Yayınları"  ; [void]$info.AppendChild($author)
    $desc     = $sdxml.CreateElement("description") ; $desc.InnerText     = "$($isim) Sözlük" ; [void]$info.AppendChild($desc)
    $email    = $sdxml.CreateElement("email")       ; [void]$email.AppendChild($sdxml.CreateWhitespace(""))    ; [void]$info.AppendChild($email)
    $website  = $sdxml.CreateElement("website")     ; [void]$website.AppendChild($sdxml.CreateWhitespace(""))  ; [void]$info.AppendChild($website)
    $date     = $sdxml.CreateElement("date")        ; $date.InnerText     = "$(Get-Date -Format "dd/MM/yyyy")" ; [void]$info.AppendChild($date)
    $dicttype = $sdxml.CreateElement("dicttype")    ; [void]$dicttype.AppendChild($sdxml.CreateWhitespace("")) ; [void]$info.AppendChild($dicttype)

    [void]$root.AppendChild($info)

    # $contents = $sdxml.CreateElement("contents")
    foreach ($i in $eklenmis) {
        $article = $sdxml.CreateElement("article")
        $key = $sdxml.CreateElement("key") ; $key.InnerText = $i.BASLIK ; [void]$article.AppendChild($key)
        if($i.ContainsKey("CEKIM") -and $i.CEKIM.Count -gt 0)
        {
            foreach ($c in $i.CEKIM)
            {
                if ($c -ine $i.BASLIK)
                {
                    $syn = $sdxml.CreateElement("synonym")
                    $syn.InnerText = $c
                    [void]$article.AppendChild($syn)
                }
            }
        }
        $cdata = $sdxml.CreateCDataSection($i.TANIM)
        $defi = $sdxml.CreateElement("definition") ; [void]$defi.SetAttribute("type", "h") ; [void]$defi.AppendChild($cdata) ; [void]$article.AppendChild($defi)
        [void]$root.AppendChild($article)
    }
    # [void]$root.AppendChild($contents)
    [void]$sdxml.AppendChild($root)
    [void]$sdxml.InsertBefore($header, $root) 
    $nobom = New-Object System.Text.UTF8Encoding $false
    $xml_fname = "$($dosyaismi)_stardict_textual.xml"
    [System.IO.TextWriter]$writer = New-Object System.IO.StreamWriter($xml_fname, $false, $nobom)
    $sdxml.Save($writer)
    $writer.Dispose()
    Write-Host "*** $($xml_fname) yazıldı." -BackgroundColor Green -ForegroundColor Black
}

if ($TSV) 
{
    $o = foreach ($i in $eklenmis) {
        "$($i.BASLIK)`t$($i.TANIM)"
    }
    $tsv_fname = "$($dosyaismi).tsv"
    $o | Out-File -FilePath $tsv_fname -Encoding UTF8
    Write-Host "*** $($tsv_fname) yazıldı." -BackgroundColor Green -ForegroundColor Black
}
