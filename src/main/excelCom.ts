import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type {
  ImportedControl,
  ImportedModule,
  ImportedProp,
  ImportJob,
  ImportResult,
  VbaJob,
  VbaReference,
  VbaRunResult,
} from '../core';

/** Excel の処理の制限時間。超えたら起動した Excel を強制終了する */
const TIMEOUT_MS = 180_000;

/**
 * 画面に出さない Excel を操作する PowerShell スクリプト（共通部分）。
 * - 自分専用の Excel を新しく起動する（利用者が開いている Excel とは混ざらない）
 * - マクロは実行しない（AutomationSecurity = ForceDisable）、確認ダイアログは出さない、イベントも止める
 * - 自動保存は切る（OneDrive 上のファイルを勝手に保存しないように）
 * - 何があっても最後に Excel を終了し、結果を JSON ファイルに書く
 */
const HEAD = String.raw`
param([Parameter(Mandatory = $true)][string]$JobPath)
$ErrorActionPreference = 'Stop'
$job = [IO.File]::ReadAllText($JobPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$result = [ordered]@{ ok = $false; error = $null; errorKind = $null }
$script:where = ''
$xl = $null
$wb = $null

function Fail([string]$kind, [string]$message) {
  $result.errorKind = $kind
  throw $message
}
`;

/** Excel を起動し、ブックを開いて VBA プロジェクトを取り出すまで（$readOnly で開き方を変える） */
const OPEN = (readOnly: boolean) =>
  String.raw`
try {
  try { $xl = New-Object -ComObject Excel.Application }
  catch { Fail 'no-excel' "デスクトップ版 Excel を起動できません（$($_.Exception.Message)）" }

  # 時間切れのときに xlCode が強制終了できるよう、起動した Excel のプロセス番号を書いておく
  try {
    Add-Type -Namespace XlCode -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
    $xlPid = [uint32]0
    [void][XlCode.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
    [IO.File]::WriteAllText($job.pidPath, [string]$xlPid)
  } catch {}

  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  $xl.AutomationSecurity = 3

  $script:where = 'ブックを開く'
  $wb = $xl.Workbooks.Open($job.input, 0, ${readOnly ? '$true' : '$false'})
  try { $wb.AutoSaveOn = $false } catch {}

  try {
    $vbp = $wb.VBProject
    if ($null -eq $vbp -or $null -eq $vbp.VBComponents) { throw 'VBProject を取得できません' }
  } catch {
    Fail 'vbom' "VBA プロジェクトにアクセスできません（$($_.Exception.Message)）"
  }

`;

const TAIL = String.raw`
} catch {
  $msg = $_.Exception.Message
  if ($script:where -and -not $result.errorKind) { $msg = "$($script:where): $msg" }
  $result.error = $msg
  if (-not $result.errorKind) { $result.errorKind = 'other' }
} finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($xl) { try { $xl.Quit() } catch {} }
  foreach ($o in @($wb, $xl)) {
    if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [IO.File]::WriteAllText($job.resultPath, ($result | ConvertTo-Json -Compress -Depth 20), (New-Object Text.UTF8Encoding $false))
}
`;

/** ビルド: VBA を書き込んで .xlsm として保存する */
export const VBA_SCRIPT =
  HEAD +
  String.raw`
function Conv($p) {
  switch ($p.type) {
    'string' { return [string]$p.value }
    'int' { return [int]$p.value }
    'double' { return [double]$p.value }
    'bool' { return [bool]$p.value }
  }
  return $p.value
}

# Font.Size のような . 区切りのプロパティにも値を入れる
function Set-Path($obj, [string]$name, $value) {
  $parts = $name.Split('.')
  $o = $obj
  for ($i = 0; $i -lt $parts.Length - 1; $i++) { $o = $o.($parts[$i]) }
  $o.($parts[$parts.Length - 1]) = $value
}

function Set-Props($obj, $props, [string]$sheet) {
  foreach ($p in @($props)) {
    if ($null -eq $p) { continue }
    $script:where = "$sheet $($p.line) 行目（$($p.name)）"
    Set-Path $obj $p.name (Conv $p)
  }
}

function Set-Code($comp, [string]$code) {
  $cm = $comp.CodeModule
  $n = $cm.CountOfLines
  if ($n -gt 0) { $cm.DeleteLines(1, $n) }
  if ($code.Length -gt 0) { $cm.AddFromString($code) }
}

function Find-Doc([string]$name) {
  foreach ($c in $vbp.VBComponents) { if ($c.Type -eq 100 -and $c.Name -ieq $name) { return $c } }
  return $null
}

# シート名から、そのシートのコードモジュールを探す
function Find-SheetDoc([string]$sheetName) {
  foreach ($c in $vbp.VBComponents) {
    if ($c.Type -ne 100 -or $c.Name -ieq $wb.CodeName) { continue }
    try { if ($c.Properties.Item('Name').Value -eq $sheetName) { return $c } } catch {}
  }
  try {
    $cn = $wb.Worksheets.Item($sheetName).CodeName
    if ($cn) { return Find-Doc $cn }
  } catch {}
  return $null
}

function Add-Controls($controls, $nodes, [string]$sheet) {
  foreach ($n in @($nodes)) {
    if ($null -eq $n) { continue }
    $script:where = "$sheet $($n.line) 行目（$($n.type) $($n.name)）"
    $ctl = $controls.Add($n.progId, $n.name, $true)
    Set-Props $ctl $n.props $sheet
    if ($n.type -eq 'MultiPage') {
      # 最初からあるページは使い回し、足りなければ足し、余れば消す
      $pages = @($n.children | Where-Object { $null -ne $_ })
      for ($i = 0; $i -lt $pages.Count; $i++) {
        $pg = $pages[$i]
        $script:where = "$sheet $($pg.line) 行目（Page $($pg.name)）"
        if ($i -lt $ctl.Pages.Count) { $page = $ctl.Pages.Item($i); $page.Name = $pg.name }
        else { $page = $ctl.Pages.Add($pg.name) }
        Set-Props $page $pg.props $sheet
        Add-Controls $page.Controls $pg.children $sheet
      }
      while ($ctl.Pages.Count -gt $pages.Count) { $ctl.Pages.Remove($ctl.Pages.Count - 1) }
    } elseif ($n.type -eq 'Frame') {
      Add-Controls $ctl.Controls $n.children $sheet
    }
  }
}

` +
  OPEN(false) +
  String.raw`
  # 参照設定（無いものだけ追加する）
  foreach ($r in @($job.references)) {
    if ($null -eq $r) { continue }
    $script:where = "参照設定 $($r.description)"
    $has = $false
    foreach ($x in @($vbp.References)) { try { if ($x.GUID -ieq $r.guid) { $has = $true } } catch {} }
    if (-not $has) {
      try { [void]$vbp.References.AddFromGuid([string]$r.guid, [int]$r.major, [int]$r.minor) }
      catch { throw "参照設定「$($r.description) $($r.guid)」を追加できません。この PC にそのライブラリが無い可能性があります（$($_.Exception.Message)）" }
    }
  }

  foreach ($m in @($job.modules)) {
    if ($null -eq $m) { continue }
    $script:where = $m.sheet
    switch ($m.kind) {
      'standard' {
        $c = $vbp.VBComponents.Add(1); $c.Name = $m.name; Set-Code $c $m.code
      }
      'class' {
        # 既にある同じ名前のシートのコード（Sheet1 など）なら、そこに書く
        $d = Find-Doc $m.name
        if ($d) { Set-Code $d $m.code }
        else { $c = $vbp.VBComponents.Add(2); $c.Name = $m.name; Set-Code $c $m.code }
      }
      'workbook' {
        $d = Find-Doc $wb.CodeName
        if (-not $d) { throw 'ブックのコードモジュール（ThisWorkbook）が見つかりません' }
        Set-Code $d $m.code
      }
      'sheet' {
        $d = Find-SheetDoc $m.targetSheet
        if (-not $d) { throw "シート「$($m.targetSheet)」のコードモジュールが見つかりません" }
        Set-Code $d $m.code
      }
      'form' {
        $c = $vbp.VBComponents.Add(3)
        $c.Name = $m.name
        foreach ($p in @($m.form.props)) {
          if ($null -eq $p) { continue }
          $script:where = "$($m.sheet) $($p.line) 行目（$($p.name)）"
          $v = Conv $p
          if ($p.name.Contains('.')) { Set-Path $c.Designer $p.name $v }
          else {
            try { $c.Properties.Item($p.name).Value = $v } catch { Set-Path $c.Designer $p.name $v }
          }
        }
        Add-Controls $c.Designer.Controls $m.form.children $m.sheet
        $script:where = $m.sheet
        Set-Code $c $m.code
      }
    }
  }

  $script:where = '.xlsm として保存'
  $wb.SaveAs($job.output, 52)
  $result.ok = $true
` +
  TAIL;

/** 取り込み: モジュール・フォームの配置・参照設定を読み出し、シートだけを .xlsx として保存する */
export const IMPORT_SCRIPT =
  HEAD +
  String.raw`
$CtlProps = @('Caption', 'Left', 'Top', 'Width', 'Height', 'TabIndex', 'TabStop', 'Visible', 'Enabled', 'Locked',
  'ControlTipText', 'ForeColor', 'BackColor', 'BackStyle', 'BorderStyle', 'BorderColor', 'SpecialEffect', 'TextAlign',
  'Alignment', 'AutoSize', 'WordWrap', 'MultiLine', 'MaxLength', 'PasswordChar', 'ScrollBars', 'Value', 'Style',
  'ColumnCount', 'ColumnWidths', 'BoundColumn', 'ColumnHeads', 'ListStyle', 'MultiSelect', 'MatchEntry', 'RowSource',
  'ControlSource', 'GroupName', 'Min', 'Max', 'SmallChange', 'LargeChange', 'Orientation', 'Default', 'Cancel',
  'Accelerator', 'PictureSizeMode', 'TabOrientation', 'EnterKeyBehavior', 'TabKeyBehavior', 'IMEMode')
$FontProps = @('Name', 'Size', 'Bold', 'Italic', 'Underline')
$FormProps = @('Caption', 'Width', 'Height', 'BackColor', 'ForeColor', 'StartUpPosition', 'ShowModal')
# 既定値と同じでも必ず書くもの
$Always = @('Caption', 'Left', 'Top', 'Width', 'Height')
$defaults = @{}
$script:tmpForm = $null
$warnings = @()

if (-not (Get-Command Get-TypeName -ErrorAction SilentlyContinue)) {
  Add-Type -AssemblyName Microsoft.VisualBasic
  function script:Get-TypeName($o) { [Microsoft.VisualBasic.Information]::TypeName($o) }
}

# 文字列・数値・真偽値だけを取り出す（それ以外や、プロパティが無いときは $null）
function Is-Value($v) {
  return ($v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [int16] -or $v -is [int64] -or
    $v -is [byte] -or $v -is [single] -or $v -is [double] -or $v -is [decimal])
}

function Read-Prop($o, [string]$n) {
  try { $v = $o.$n } catch { return $null }
  if (Is-Value $v) { return $v }
  return $null
}

function Has-Picture($o) {
  try { $p = $o.Picture; return ($null -ne $p -and [int64]$p.Handle -ne 0) } catch { return $false }
}

# 種類ごとの既定値（一時的なフォームに同じ種類のコントロールを置いて読む）
function Get-Defaults([string]$type) {
  if ($defaults.ContainsKey($type)) { return $defaults[$type] }
  $d = @{}
  try {
    if (-not $script:tmpForm) { $script:tmpForm = $vbp.VBComponents.Add(3) }
    $c = $script:tmpForm.Designer.Controls.Add("Forms.$type.1", "xlcodeDefault$type", $false)
    foreach ($n in $CtlProps) { $v = Read-Prop $c $n; if ($null -ne $v) { $d[$n] = $v } }
    $f = $null; try { $f = $c.Font } catch {}
    if ($f) { foreach ($n in $FontProps) { $v = Read-Prop $f $n; if ($null -ne $v) { $d["Font.$n"] = $v } } }
  } catch {}
  $defaults[$type] = $d
  return $d
}

function Collect-Props($o, [string]$type) {
  $def = if ($type -eq 'Page') { @{} } else { Get-Defaults $type }
  $out = @()
  foreach ($n in $CtlProps) {
    $v = Read-Prop $o $n
    if ($null -eq $v) { continue }
    if (-not ($Always -contains $n) -and $def.ContainsKey($n) -and $def[$n] -ceq $v) { continue }
    $out += [ordered]@{ name = $n; value = $v }
  }
  $f = $null; try { $f = $o.Font } catch {}
  if ($f) {
    foreach ($n in $FontProps) {
      $v = Read-Prop $f $n
      if ($null -eq $v) { continue }
      if ($def.ContainsKey("Font.$n") -and $def["Font.$n"] -ceq $v) { continue }
      $out += [ordered]@{ name = "Font.$n"; value = $v }
    }
  }
  return ,$out
}
` +
  OPEN(true) +
  String.raw`
  $mods = @()
  foreach ($c in @($vbp.VBComponents)) {
    $script:where = [string]$c.Name
    $m = [ordered]@{ name = [string]$c.Name; type = [int]$c.Type; code = @() }
    $n = $c.CodeModule.CountOfLines
    if ($n -gt 0) { $m.code = @(([string]$c.CodeModule.Lines(1, $n)) -split "\r?\n") }
    if ($c.Type -eq 100) {
      if ($c.Name -ieq $wb.CodeName) { $m.workbook = $true }
      else { try { $m.sheet = [string]$c.Properties.Item('Name').Value } catch {} }
    }
    if ($c.Type -eq 3) {
      $form = [ordered]@{ props = @(); controls = @() }
      foreach ($pn in $FormProps) {
        try { $v = $c.Properties.Item($pn).Value } catch { continue }
        if (Is-Value $v) { $form.props += [ordered]@{ name = $pn; value = $v } }
      }
      if (Has-Picture $c.Designer) { $form.picture = $true }
      foreach ($ctl in @($c.Designer.Controls)) {
        $t = [string](Get-TypeName $ctl)
        $e = [ordered]@{ type = $t; name = [string]$ctl.Name; parent = ''; props = @() }
        try { $e.parent = [string]$ctl.Parent.Name } catch {}
        $e.props = Collect-Props $ctl $t
        if (Has-Picture $ctl) { $e.picture = $true }
        $form.controls += $e
        if ($t -eq 'MultiPage') {
          foreach ($pg in @($ctl.Pages)) {
            $form.controls += [ordered]@{ type = 'Page'; name = [string]$pg.Name; parent = [string]$ctl.Name; props = (Collect-Props $pg 'Page') }
          }
        }
      }
      $m.form = $form
    }
    $mods += $m
  }

  $refs = @()
  foreach ($r in @($vbp.References)) {
    try {
      if ($r.BuiltIn) { continue }
      if ([int]$r.Type -ne 0) { $warnings += "参照設定「$($r.Name)」はライブラリではない（ほかのブックへの参照）ため取り込めません"; continue }
      $refs += [ordered]@{ guid = [string]$r.GUID; major = [int]$r.Major; minor = [int]$r.Minor; description = [string]$r.Description; name = [string]$r.Name }
    } catch {}
  }

  if ($script:tmpForm) { try { $vbp.VBComponents.Remove($script:tmpForm) } catch {} }
  $script:where = 'シートを .xlsx として保存'
  $wb.SaveAs($job.output, 51)
  $result.modules = $mods
  $result.references = $refs
  $result.warnings = $warnings
  $result.ok = $true
` +
  TAIL;

const UTF8_BOM = Buffer.from([0xef, 0xbb, 0xbf]);

function run(file: string, args: string[], timeout: number): Promise<{ timedOut: boolean; error?: string }> {
  return new Promise((resolve) => {
    execFile(
      file,
      args,
      { windowsHide: true, timeout, killSignal: 'SIGKILL', maxBuffer: 4 * 1024 * 1024 },
      (e, _out, err) => {
        if (!e) return resolve({ timedOut: false });
        const timedOut = (e as NodeJS.ErrnoException & { killed?: boolean }).killed === true;
        resolve({ timedOut, error: String(err || e.message).trim() });
      },
    );
  });
}

export interface RunOptions {
  timeout?: number;
  /** テスト用: PowerShell の実行ファイルとスクリプト */
  exe?: string;
  script?: string;
}

type RawResult = { ok: boolean; error: string | null; errorKind: string | null } & Record<string, unknown>;

/** スクリプトを実行し、結果の JSON を返す */
async function runScript(name: string, defaultScript: string, job: object, opts: RunOptions): Promise<RawResult> {
  const timeout = opts.timeout ?? TIMEOUT_MS;
  const dir = await mkdtemp(path.join(tmpdir(), 'xlcode-vba-'));
  try {
    const script = path.join(dir, `${name}.ps1`);
    const jobPath = path.join(dir, 'job.json');
    const resultPath = path.join(dir, 'result.json');
    const pidPath = path.join(dir, 'excel.pid');
    // Windows PowerShell 5.1 は BOM の無いスクリプトを Shift_JIS として読むため BOM を付ける
    await writeFile(
      script,
      Buffer.concat([UTF8_BOM, Buffer.from((opts.script ?? defaultScript).replace(/\r?\n/g, '\r\n'), 'utf8')]),
    );
    await writeFile(jobPath, JSON.stringify({ ...job, resultPath, pidPath }), 'utf8');
    const r = await run(
      opts.exe ?? 'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-JobPath', jobPath],
      timeout,
    );
    if (r.timedOut) {
      const pid = Number((await readFile(pidPath, 'utf8').catch(() => '')).trim());
      if (pid > 0) await run('taskkill.exe', ['/PID', String(pid), '/T', '/F'], 20_000);
      return {
        ok: false,
        errorKind: 'timeout',
        error:
          `Excel の処理が ${Math.round(timeout / 1000)} 秒で終わらなかったため中止しました` +
          (pid > 0
            ? '（起動した Excel は終了させました）'
            : '（画面に出ない Excel が残っている場合は、タスク マネージャーで終了してください）'),
      };
    }
    const text = await readFile(resultPath, 'utf8').catch(() => null);
    if (text === null) {
      return { ok: false, errorKind: 'other', error: `PowerShell を実行できません${r.error ? `: ${r.error}` : ''}` };
    }
    return JSON.parse(text.replace(/^\uFEFF/, '')) as RawResult;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function failure(res: RawResult): VbaRunResult {
  return {
    ok: false,
    error: res.error ?? '原因不明',
    errorKind: (res.errorKind as VbaRunResult['errorKind']) ?? 'other',
  };
}

/** VBA の書き込みを Excel に依頼する（Windows のみ） */
export async function runVbaJob(job: VbaJob, opts: RunOptions = {}): Promise<VbaRunResult> {
  const res = await runScript('build', VBA_SCRIPT, job, opts);
  return res.ok ? { ok: true } : failure(res);
}

/** PowerShell の JSON では要素が 1 つの配列が値になることがあるため、配列にそろえる */
function list<T>(v: unknown): T[] {
  if (v === null || v === undefined) return [];
  return (Array.isArray(v) ? v : [v]) as T[];
}

/** 既存のツールの読み取りを Excel に依頼する（Windows のみ） */
export async function runImportJob(job: ImportJob, opts: RunOptions = {}): Promise<ImportResult> {
  const res = await runScript('import', IMPORT_SCRIPT, job, opts);
  if (!res.ok) return failure(res);
  const modules = list<ImportedModule>(res.modules).map((m) => ({
    ...m,
    code: list<string>(m.code).map(String),
    form: m.form && {
      ...m.form,
      props: list<ImportedProp>(m.form.props),
      controls: list<ImportedControl>(m.form.controls).map((c) => ({ ...c, props: list<ImportedProp>(c.props) })),
    },
  }));
  return {
    ok: true,
    modules,
    references: list<VbaReference>(res.references),
    warnings: list<string>(res.warnings),
  };
}
