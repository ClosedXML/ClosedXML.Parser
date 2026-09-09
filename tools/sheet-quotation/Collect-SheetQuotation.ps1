param(
    [Parameter(Mandatory)][string]$cpFile,
    [Parameter(Mandatory)][string]$out,
    [Parameter(Mandatory)][string]$namesOut,
    # first: <char>tag. mid: tag<char>zz. last: tag<char>.
    # `mid` and `last` both feed the "next" (non-first) table; collect both and union
    # them, because the end of a name is a position Excel can treat specially.
    [ValidateSet('first','mid','last')][string]$mode = 'mid'
)
$ErrorActionPreference = 'Stop'

# Sheet names must be unique and must not look like a cell reference, or Excel quotes
# them for that reason instead of for the character under test. The 5-letter tag is past
# the 3-letter column limit, so no name here can be mistaken for a reference.
function Enc([int]$n) { $s=''; for ($k=0;$k -lt 4;$k++) { $s = [char](97 + ($n % 26)) + $s; $n = [int][math]::Floor($n/26) }; return $s }

$cps = [int[]](Get-Content $cpFile | Where-Object { $_.Trim() } | ForEach-Object { [int]("0x$_") })

# Workbook.SaveAs resolves a relative path against Excel's working directory, not this
# session's, so a relative -out would be deleted here and written somewhere else - and the
# harvest step would then read a stale workbook and record the wrong answers, silently.
$out = [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location).Path, $out))
$namesOut = [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location).Path, $namesOut))
if (Test-Path $out) { Remove-Item $out -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$rows = New-Object System.Collections.Generic.List[string]
try {
    $wb = $xl.Workbooks.Add()
    $probe = $wb.Worksheets.Item(1)
    $probe.Name = 'Probe'
    $row = 0
    for ($i = 0; $i -lt $cps.Count; $i++) {
        $cp = $cps[$i]; $ch = [char]$cp; $tag = 'q' + (Enc $i)
        $name = switch ($mode) {
            'first' { "$ch$tag" }
            'mid'   { "$tag$ch" + 'zz' }
            'last'  { "$tag$ch" }
        }
        $ws = $null
        try { $ws = $wb.Worksheets.Add(); $ws.Name = $name }
        catch { $rows.Add(("{0:X4},INVALID,," -f $cp)); continue }
        $row++
        # U+0027 is legal mid-name, so the reference written here has to escape it, or
        # the probe formula is malformed and Excel rejects the assignment.
        $ref = $name.Replace("'", "''")
        try { $probe.Range("A$row").Formula = "='$ref'!A1" }
        catch { $row--; $rows.Add(("{0:X4},NOFORMULA,," -f $cp)); continue }
        $rows.Add(("{0:X4},OK,{1},{2}" -f $cp, $row, $name))
    }
    $wb.SaveAs($out, 51)
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
[IO.File]::WriteAllLines($namesOut, $rows, [Text.UTF8Encoding]::new($false))
"built $out ($($rows.Count) entries, $row formulas)"
