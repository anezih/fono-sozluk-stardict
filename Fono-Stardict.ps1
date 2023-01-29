#Requires -Version 7

[CmdletBinding()]
param 
(
    [switch]$GLS,
    [switch]$Textual,
    [switch]$TSV,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("EN","FR","ES","DE","RU")]
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
        $tanim_html = $tanim_html.Replace("color:#808080;","color:#000080;").
                                Replace("color:#C0C0C0;","color:#008000;").
                                Replace("font-size:9pt;","").
                                Replace("font-size:12pt;","").
                                Replace("font-size:10pt;","").
                                Replace("font-family:&quot;Arial TUR&quot;, sans-serif;","").
                                Replace("color:#0000FF","color:#000000").
                                Replace('<em style="color:#9999FF;">','<em style="color:#A0522D;">').
                                Replace(' style=""',"")
        @{
            BASLIK = $baslik.ReplaceLineEndings("").Trim()
            TANIM  = $tanim_html.ReplaceLineEndings("").Trim()
        }
    }
    return $cikti
}

function FRTR
{
    param(
        [string]$fileName,
        [int]$startOffset
    )
    Add-Type -AssemblyName $PSScriptRoot\Modul\RtfPipe.dll
    
    $byteArray = [System.IO.File]::ReadAllBytes($fileName)
    $byteArray = $byteArray[$startOffset..$byteArray.Length]
    [System.Collections.ArrayList]$arr = @()
    [System.Collections.ArrayList]$offsets = @()
    $enc = [Text.Encoding]::UTF8

    [void]$offsets.Add(0)
    for(($i = 0); ($i -lt $byteArray.Length);($i++))
    {
        if($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0xDA -and $byteArray[($i-3)..($i-1)] -eq 0) # tek başına 32869,32763
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x9C -and $byteArray[($i-3)..($i-1)] -eq 0)
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x01 -and $byteArray[($i-3)..($i-1)] -eq 0) # ikisini ekleyince 33129,32778
        {
            [void]$offsets.Add($i)
        }
        elseif($byteArray[$i] -eq 0x78 -and $byteArray[$i+1] -eq 0x5E -and $byteArray[($i-3)..($i-1)] -eq 0) # bunu da ekleyince 33408,32778
        {
            [void]$offsets.Add($i)
        }
    }

    # $cnt = 1

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
            
            $rtf = $enc.GetString($out[4..($out.Length-3)])
            # if(-not $rtf.EndsWith("}")) { continue }
            $html = [RtfPipe.Rtf]::ToHtml([System.String]$rtf)
            if($html -match "([\w+'-]+(?:\s)\w+)</strong><strong>(\w+)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "<strong>(\w+(?:\s)\w+)</strong><strong[^>]+>(\w+)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "<strong>(\w+)</strong><strong[^>]+>(\w+)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "<strong>(\w+)</strong><strong[^>]+>(\w+(?:\s)\w+)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "<strong>(\w+)</strong><strong[^>]+>(.*?)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "(\w+)</strong><strong>(\w+)</strong>")
            {
                $hw = $Matches.1 + $Matches.2
            }
            elseif($html -match "<strong[^>]+>(.*?)</strong>")
            {
                $hw = $Matches.1
            }
            elseif($html -match "<strong> </strong><strong><em>(.*?)</em></strong>")
            {
                $hw = $Matches.1
            }
            elseif($html -match "<strong>(.*?)</strong>")
            {
                $hw = $Matches.1
            }
            else
            {
                continue
            }
            $html = $html -replace 
                    "<p[^>]+><br></p>", "" -replace
                    "<div[^>]+>", "<div>" -replace
                    "<p[^>]+>", "<p>" -replace
                    # "<strong><em[^>]*>", '<strong><em style="color:#0077B3;">' -replace
                    # "[^(?:<strong>)]<em[^>]+>", ' <em style="color:#CC7A00;">' -replace
                    # "</span><em[^>]*>", '</span><em style="color:#CC7A00;">' -replace
                    # "</strong><em[^>]*>", '</strong><em style="color:#CC7A00;">' -replace
                    # "[^(?:<strong>)]<em>",' <em style="color:#CC7A00;">' -replace
                    # "[^(?:<strong>)](?:</span>)<em>", ' </span><em style="color:#CC7A00;">' -replace
                    "<strong[^>]+>", "<strong>" -replace
                    "<span[^>]+>", "<span>" -replace
                    "<em[^>]+>", "<em>"
            $temiz = ""
            foreach ($c in $html.ToCharArray()) {
                if (-not [char]::IsControl($c)) {
                    $temiz += $c
                }
            }
            [void]$arr.Add(@{ BASLIK = $hw.ReplaceLineEndings("").Trim() ; TANIM = $temiz.ReplaceLineEndings("").Trim()})
            # if($cnt % 5000 -eq 0){ Write-Host "$($fileName)'den $($cnt) numaralı girdi eklendi" } #  
            # $cnt++
        }
        catch
        {
            # "$($_) $($rtf)" >> exceptions.txt
        }
        finally
        {
            $inputStream.Dispose() ; $outputStream.Dispose() ; $compressionStream.Dispose()
        }
    }
    return $arr
}

if (-not ($GLS -or $TSV -or $Textual)) {
    Write-Host "En az bir tane çıktı türü belirtmelisiniz."
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
        $trye_  = ENTR $ING_TR
        if ($Hunspell) {
            Import-Module $PSScriptRoot\Modul\HunspellWordForms.psm1
            $sozluk_tr = New-HunspellDictionaryObject -path $PSScriptRoot\Hunspell\tr_TR.dic
            $sozluk_en = New-HunspellDictionaryObject -path $PSScriptRoot\Hunspell\en_US.dic
            $syc0 = 0
            $toplam0 = $trden_.Count
            $trden = foreach ($i in $trden_) {
                $cekim = (Get-WordForms -Dictionary $sozluk_tr -Word $i.BASLIK -NoPFX -NoCross).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
                $syc0++
                if($syc0 % 100 -eq 0) {
                    Write-Progress -Activity "TR son ek bulma: $("{0:N0}" -f $syc0) / $("{0:N0}" -f $toplam0)" -Status "$("{0:P}" -f  $($syc0/$toplam0))" -PercentComplete (($syc0/$toplam0)*100)
                }
            }
            $syc1 = 0
            $toplam1 = $trye_.Count
            $trye = foreach ($i in $trye_) {
                $cekim = (Get-WordForms -Dictionary $sozluk_en -Word $i.BASLIK -NoPFX -NoCross).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
                $syc1++
                if($syc1 % 100 -eq 0) {
                    Write-Progress -Activity "EN son ek bulma: $("{0:N0}" -f $syc1) / $("{0:N0}" -f $toplam1)" -Status "$("{0:P}" -f  $($syc1/$toplam1))" -PercentComplete (($syc1/$toplam1)*100)
                }
            }
        }
    }

    "FR"
    {
        $isim      = "Fono Fransızca⇄Türkçe"
        $dosyaismi = "fono_fr_tr"

        $trden_ = FRTR -fileName $PSScriptRoot\SozlukDosyalari\TURFRE_P.KDD -startOffset 795999
        $trye_  = FRTR -fileName $PSScriptRoot\SozlukDosyalari\FRETUR_P.KDD -startOffset 1086923
        if ($Hunspell) {
            Import-Module $PSScriptRoot\Modul\HunspellWordForms.psm1
            $sozluk_tr = New-HunspellDictionaryObject -path $PSScriptRoot\Hunspell\tr_TR.dic
            $sozluk_fr = New-HunspellDictionaryObject -path $PSScriptRoot\Hunspell\fr_FR.dic
            $syc2 = 0
            $toplam2 = $trden_.Count
            $trden = foreach ($i in $trden_) {
                $cekim = (Get-WordForms -Dictionary $sozluk_tr -Word $i.BASLIK -NoPFX -NoCross).SFX
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
                $syc2++
                if($syc2 % 100 -eq 0) {
                    Write-Progress -Activity "TR son ek bulma: $("{0:N0}" -f $syc2) / $("{0:N0}" -f $toplam2)" -Status "$("{0:P}" -f  $($syc2/$toplam2))" -PercentComplete (($syc2/$toplam2)*100)
                }
            }
            $syc3 = 0
            $toplam3 = $trye_.Count
            $trye = foreach ($i in $trye_) {
                if($i.BASLIK.Contains(","))
                {
                    $sozcuk, $cinsiyet = $i.BASLIK.Split(",") ; $cinsiyet = $cinsiyet.Trim()
                    $cekim  = (Get-WordForms -Dictionary $sozluk_fr -Word $sozcuk -NoPFX -NoCross).SFX
                    $cekim += (Get-WordForms -Dictionary $sozluk_fr -Word ($sozcuk+$cinsiyet) -NoPFX -NoCross).SFX
                    if($sozcuk -notin $cekim) {$cekim += $sozcuk}
                }
                else
                {
                    $cekim = (Get-WordForms -Dictionary $sozluk_fr -Word $i.BASLIK -NoPFX -NoCross).SFX
                }
                @{
                    BASLIK = $i.BASLIK
                    TANIM  = $i.TANIM
                    CEKIM  = $cekim
                }
                $syc3++
                if($syc3 % 100 -eq 0) {
                    Write-Progress -Activity "FR son ek bulma: $("{0:N0}" -f $syc3) / $("{0:N0}" -f $toplam3)" -Status "$("{0:P}" -f  $($syc3/$toplam3))" -PercentComplete (($syc3/$toplam3)*100)
                }
            }
        }
    }

    { @("ES","DE","RU") -contains $_ }
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
            $a = foreach($j in [System.Collections.Generic.HashSet[string]]($i.CEKIM)){
                if ($j -and ($j -notmatch "\s+") -and ($j -ne $i.BASLIK)) {
                    $j
                }
            }
            $b  = $i.BASLIK + "|"
            $b += $a -join "|"
            if($b.EndsWith("|")) { $b = $b.Substring(0, $b.Length-1) }
            "$($b)`n$($i.TANIM)`n`n"
        }
        else
        {
            "$($i.BASLIK)`n$($i.TANIM)`n`n"
        }
    }   
    # $o = $ustbilgi + $eklenmis -> nedense her maddeden önce bir boşluk ekliyor
    $o | Out-File -FilePath "$($dosyaismi).gls" -Encoding UTF8 -NoNewline -Append
}
if ($Textual) {
    [xml]$sdxml = New-Object System.Xml.XmlDocument

    [System.Xml.XmlNode]$header = $sdxml.CreateXmlDeclaration("1.0", "UTF-8", $null)   
    # [System.Xml.XmlAttribute]$namespace = $sdxml.CreateAttribute("xmlns", "xi", "http://www.w3.org/2003/XInclude")
    
    $root = $sdxml.CreateElement("stardict")
    # $root.Attributes.Append($namespace)

    $info = $sdxml.CreateElement("info")
    $version = $sdxml.CreateElement("version") ; $version.InnerText = "3.0.0" ; [void]$info.AppendChild($version)
    $bookname = $sdxml.CreateElement("bookname") ; $bookname.InnerText = "$($isim) Sözlük" ; [void]$info.AppendChild($bookname)
    $author = $sdxml.CreateElement("author") ; $author.InnerText = "Fono Yayınları" ; [void]$info.AppendChild($author)
    $desc = $sdxml.CreateElement("description") ; $desc.InnerText = "$($isim) Sözlük" ; [void]$info.AppendChild($desc)
    $email = $sdxml.CreateElement("email") ; [void]$email.AppendChild($sdxml.CreateWhitespace("")) ; [void]$info.AppendChild($email)
    $website = $sdxml.CreateElement("website") ; [void]$website.AppendChild($sdxml.CreateWhitespace("")) ; [void]$info.AppendChild($website)
    $date = $sdxml.CreateElement("date") ; $date.InnerText = "$(Get-Date -Format "dd/MM/yyyy")" ; [void]$info.AppendChild($date)
    $dicttype = $sdxml.CreateElement("dicttype") ; [void]$dicttype.AppendChild($sdxml.CreateWhitespace("")) ; [void]$info.AppendChild($dicttype)

    [void]$root.AppendChild($info)

    # $contents = $sdxml.CreateElement("contents")
    foreach ($i in $eklenmis) {
        $article = $sdxml.CreateElement("article")
        $key = $sdxml.CreateElement("key") ; $key.InnerText = $i.BASLIK ; [void]$article.AppendChild($key)
        if($i.ContainsKey("CEKIM") -and $i.CEKIM.Count -gt 0)
        {
            $bas = $i.BASLIK
            $a = [System.Collections.Generic.HashSet[string]]($i.CEKIM)
            foreach ($j in $a) {
                if($j -and ($j -notmatch "\s+") -and ($j -ne $bas)){
                    $syn = $sdxml.CreateElement("synonym")
                    $syn.InnerText = $j
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
    [System.IO.TextWriter]$writer = New-Object System.IO.StreamWriter("$($dosyaismi)_stardict_textual.xml", $false, $nobom)
    $sdxml.Save($writer)
    $writer.Dispose()
}

if ($TSV) 
{
    $o = foreach ($i in $eklenmis) {
        "$($i.BASLIK)`t$($i.TANIM)"
    }
    $o | Out-File -FilePath "$($dosyaismi).tsv" -Encoding UTF8
}