# Adapted from: https://gist.github.com/aarondandy/aaa622afeeb0cb86b0d4efe697c23be5
Add-Type -AssemblyName $PSScriptRoot\WeCantSpell.Hunspell.dll
function New-HunspellDictionaryObject {
    param (
        [string]$path
    )
    return [WeCantSpell.Hunspell.WordList]::CreateFromFiles($path)
}

function AppendPrefix
{
    param(
        $prefix,
        [string]$word
    )
    if($prefix.Conditions.IsStartingMatch($word) -and $word.StartsWith($prefix.Strip))
    {
        $combined = $prefix.Append + $word.Substring($prefix.Strip.Length)
        return $combined
    }
    return $false
}
function AppendSuffix
{
    param(
        $suffix,
        [string]$word
    )
    if($suffix.Conditions.IsEndingMatch($word) -and $word.EndsWith($suffix.Strip))
    {
        $combined = $word.Substring(0, $word.Length - $suffix.Strip.Length) + $suffix.Append
        return $combined
    }
    return $false
}

function Get-WordForms 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Dictionary,
        [Parameter(Mandatory=$true)]
        $Word,
        [switch]
        $NoPFX,
        [switch]
        $NoSFX,
        [switch] 
        $NoCross
    )
    if(-not $Dictionary -or -not $Word) { return }
    try 
    {
        # $allPrefixes = $Dictionary.Affix.Prefixes | ? {$Dictionary.Item($Word).ContainsFlag($_.AFlag)}
        # $allSuffixes = $Dictionary.Affix.Suffixes | ? {$Dictionary.Item($Word).ContainsFlag($_.AFlag)}
        $item = $Dictionary.Item($Word)
        $allPrefixes = foreach ($p in $Dictionary.Affix.Prefixes) {
            if($item.ContainsFlag($p.AFlag)) { $p }
        }
        $allSuffixes = foreach ($s in $Dictionary.Affix.Suffixes) {
            if($item.ContainsFlag($s.AFlag)) { $s }
        }

    }

    catch { return }

    if(-not $NoPFX)
    {
        $allPrefixesE = foreach($p in $allPrefixes) { $p.Entries }
        # ($allPrefixes | % {$_.Entries})
        [System.Collections.Generic.HashSet[string]]$wp = foreach($p in $allPrefixesE)
        {
            $out = AppendPrefix -prefix $p -word $Word
            if($out)
            {
                $out
            }
        }
    }

    if(-not $NoSFX)
    {
        $allSuffixesE = foreach($s in $allSuffixes) { $s.Entries }
        # ($allSuffixes | % {$_.Entries})
        [System.Collections.Generic.HashSet[string]]$ws = foreach($s in $allSuffixesE)
        {
            $out = AppendSuffix -suffix $s -word $Word
            if($out)
            {
                $out
            }
        }
    }
    
    if(-not $NoCross)
    {
        $flag = [WeCantSpell.Hunspell.AffixEntryOptions]::CrossProduct
        $allPrefixesACE = foreach($pac in $allPrefixes){
            if (($pac.Options -and $flag) -eq $flag) {
                $pac.Entries
            }
        }
        # ($allPrefixes | ? {AllowCross -value $_.Options}) | % {$_.Entries}
        [System.Collections.Generic.HashSet[string]]$wc = foreach($p in $allPrefixesACE)
        {
            $withPrefix = AppendPrefix -prefix $p -word $Word
            if($withPrefix)
            {
                $allSuffixesACE = foreach($sac in $allSuffixes){
                    if(($sac.Options -and $flag) -eq $flag) {
                        $sac.Entries
                    }
                }
                # ($allSuffixes | ? {AllowCross -value $_.Options}) | % {$_.Entries}
                foreach($s in  $allSuffixesACE)
                {
                    $crossOut = AppendSuffix -suffix $s -word $withPrefix
                    if($crossOut)
                    {
                        $crossOut
                    }
                }
            }
        }
    }

    [PSCustomObject]::new(@{
        PFX = $wp
        SFX = $ws
        Cross = $wc
    })
}

function Unmunch 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Dictionary,
        [switch]
        $NoPFX,
        [switch]
        $NoSFX,
        [switch] 
        $NoCross,
        [switch]
        $ToJson,
        [string]
        $OutPath
    )
    if(-not $Dictionary) { return }
    $all_words = [array]$Dictionary.RootWords
    $totalWords = $all_words.Length
    $cnt = 0
    $t0 = Get-Date
    $all_forms = foreach($w in $all_words) {
        $forms = Get-WordForms -Dictionary $Dictionary -Word $w -NoPFX:$NoPFX -NoSFX:$NoSFX -NoCross:$NoCross
        $fields = @{
            PFX   = $forms.PFX
            SFX   = $forms.SFX
            CROSS = $forms.CROSS
        }
        @{ $w = $fields }
        $cnt += 1
        $t1 = ((Get-Date) - $t0).TotalSeconds
        $remaining    = $totalWords - $cnt
        $avgPerItem   = $t1 / $cnt
        $secRemaining = [math]::Round(($remaining * $avgPerItem) / 5) * 5
        # Write-Host "Progress: $("{0:N0}" -f $cnt) / $("{0:N0}" -f $totalWords) | $("{0:P2}" -f ($cnt/$totalWords))`r" -NoNewLine
        Write-Progress -Activity "Progress: $("{0:N0}" -f $cnt) / $("{0:N0}" -f $totalWords)" -Status "$("{0:P}" -f  $($cnt/$totalWords))" -PercentComplete (($cnt/$totalWords)*100) -SecondsRemaining $secRemaining
    }
    $all_forms_sorted = $all_forms | Sort-Object {$_.Keys}
    if($ToJson -and $OutPath)
    {
        $all_forms_sorted | ConvertTo-Json -Depth 3 | Out-File -FilePath "$($OutPath).json" -Encoding utf8
    }
    elseif ($ToJson) {
        $all_forms_sorted | ConvertTo-Json -Depth 3 | Out-File -FilePath "unmunched.json" -Encoding utf8
    }
    else { return $all_forms_sorted }

}    
Export-ModuleMember -Function Get-WordForms, New-HunspellDictionaryObject