# テスト用の偽の Excel（COM）。build.ps1 の New-Object -ComObject Excel.Application の代わりに使う。
# 保存（SaveAs）すると、VBA プロジェクトの中身を JSON で書き出す。

class FakeCodeModule {
  [System.Collections.Generic.List[string]]$Buf = [System.Collections.Generic.List[string]]::new()
  [int]$CountOfLines = 0
  [void] DeleteLines([int]$start, [int]$count) {
    $this.Buf.RemoveRange($start - 1, $count); $this.CountOfLines = $this.Buf.Count
  }
  [void] AddFromString([string]$code) {
    foreach ($l in ($code -split "`r`n")) { $this.Buf.Add($l) }
    $this.CountOfLines = $this.Buf.Count
  }
  [string] Lines([int]$start, [int]$count) { return ($this.Buf.GetRange($start - 1, $count) -join "`r`n") }
}

class FakeProperties {
  [hashtable]$Map = @{}
  [hashtable] Item([string]$name) {
    if (-not $this.Map.ContainsKey($name)) { throw "プロパティ $name はありません" }
    return $this.Map[$name]
  }
}

class FakeControls : System.Collections.IEnumerable {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [System.Collections.IEnumerator] GetEnumerator() { return $this.List.GetEnumerator() }
  [hashtable] Add([string]$progId, [string]$name, [bool]$visible) {
    if (-not $progId.StartsWith('Forms.')) { throw "ProgID が不正です: $progId" }
    # 既定値（取り込みで「既定値と同じなら書かない」を確かめるため）
    $c = @{ ProgId = $progId; Name = $name; Font = @{ Name = 'MS UI Gothic'; Size = 9; Bold = $false }
      ForeColor = -2147483630; BackColor = -2147483633; Left = 0; Top = 0; Width = 72; Height = 24 }
    if ($progId -eq 'Forms.Frame.1') { $c.Controls = [FakeControls]::new() }
    if ($progId -eq 'Forms.MultiPage.1') {
      $c.Pages = [FakePages]::new()
      [void]$c.Pages.Add('Page1'); [void]$c.Pages.Add('Page2')
    }
    [void]$this.List.Add($c)
    return $c
  }
}

class FakePages : System.Collections.IEnumerable {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [System.Collections.IEnumerator] GetEnumerator() { return $this.List.GetEnumerator() }
  [int]$Count = 0
  [hashtable] Item([int]$i) { return $this.List[$i] }
  [hashtable] Add([string]$name) {
    $p = @{ Name = $name; Controls = [FakeControls]::new(); Font = @{} }
    [void]$this.List.Add($p)
    $this.Count = $this.List.Count
    return $p
  }
  [void] Remove([int]$i) { $this.List.RemoveAt($i); $this.Count = $this.List.Count }
}

class FakeComponent {
  [string]$Name
  [int]$Type
  [FakeCodeModule]$CodeModule = [FakeCodeModule]::new()
  [FakeProperties]$Properties = [FakeProperties]::new()
  $Designer = $null
}

class FakeComponents : System.Collections.IEnumerable {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [int]$Count = 0
  [System.Collections.IEnumerator] GetEnumerator() { return $this.List.GetEnumerator() }
  [void] Push([FakeComponent]$c) { [void]$this.List.Add($c); $this.Count = $this.List.Count }
  [void] Remove([FakeComponent]$c) { $this.List.Remove($c); $this.Count = $this.List.Count }
  [FakeComponent] Add([int]$type) {
    $c = [FakeComponent]::new()
    $c.Type = $type
    $prefix = @{ 1 = 'Module'; 2 = 'Class'; 3 = 'UserForm' }[$type]
    $n = 1
    while ($this.List | Where-Object { $_.Name -eq "$prefix$n" }) { $n++ }
    $c.Name = "$prefix$n"
    if ($type -eq 3) {
      $c.Designer = @{ Controls = [FakeControls]::new(); Font = @{} }
      foreach ($k in 'Caption', 'Width', 'Height') { $c.Properties.Map[$k] = @{ Value = $null } }
    }
    $this.Push($c)
    return $c
  }
}

class FakeReferences : System.Collections.IEnumerable {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [System.Collections.IEnumerator] GetEnumerator() { return $this.List.GetEnumerator() }
  [hashtable] AddFromGuid([string]$guid, [int]$major, [int]$minor) {
    if ($guid -like '{00000000*') { throw 'ライブラリが登録されていません' }
    $r = @{ GUID = $guid; Major = $major; Minor = $minor; BuiltIn = $false; Type = 0; Name = 'Lib'; Description = 'Lib' }
    [void]$this.List.Add($r)
    return $r
  }
}

class FakeProject {
  [FakeComponents]$VBComponents = [FakeComponents]::new()
  [FakeReferences]$References = [FakeReferences]::new()
}

class FakeSheet {
  [string]$Name
  [string]$CodeName
}

class FakeSheets {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [FakeSheet] Item([string]$name) {
    foreach ($s in $this.List) { if ($s.Name -eq $name) { return $s } }
    throw "シート $name はありません"
  }
}

class FakeWorkbook {
  [string]$CodeName = 'ThisWorkbook'
  [bool]$AutoSaveOn = $true
  [FakeSheets]$Worksheets = [FakeSheets]::new()
  hidden [FakeProject]$Project = [FakeProject]::new()
  [string]$SavedAs
  [int]$SavedFormat
  [bool]$Closed
  [string]$Path

  [void] SaveAs([string]$path, [int]$format) {
    $this.SavedAs = $path
    $this.SavedFormat = $format
    if ($format -eq 51) { Copy-Item -LiteralPath $this.Path -Destination $path; return }
    $dump = [ordered]@{ format = $format; autoSave = $this.AutoSaveOn; components = @() }
    foreach ($c in $this.Project.VBComponents.List) {
      $props = [ordered]@{}
      foreach ($k in $c.Properties.Map.Keys) { $props[$k] = $c.Properties.Map[$k].Value }
      $dump.components += [ordered]@{
        name = $c.Name; type = $c.Type; code = @($c.CodeModule.Buf); props = $props; designer = $c.Designer
      }
    }
    $dump.references = @($this.Project.References.List)
    [IO.File]::WriteAllText($path, ($dump | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
  }
  [void] Close([bool]$save) { $this.Closed = $true }
}

class FakeWorkbooks {
  [FakeWorkbook]$Last
  [FakeWorkbook] Open([string]$path, $updateLinks, $readOnly) {
    if (-not (Test-Path -LiteralPath $path)) { throw "ファイルがありません: $path" }
    if ($env:FAKE_EXCEL_SLEEP) { Start-Sleep -Seconds ([int]$env:FAKE_EXCEL_SLEEP) }
    $wb = [FakeWorkbook]::new()
    $wb.Path = $path
    $refs = $wb.Project.References.List
    [void]$refs.Add(@{ GUID = '{000204EF-0000-0000-C000-000000000046}'; Major = 4; Minor = 2; BuiltIn = $true; Type = 0; Name = 'VBA'; Description = 'VBA' })
    $comps = $wb.Project.VBComponents
    $doc = [FakeComponent]::new(); $doc.Name = 'ThisWorkbook'; $doc.Type = 100
    $doc.Properties.Map['Name'] = @{ Value = 'input.xlsx' }
    $comps.Push($doc)
    $i = 1
    foreach ($s in ($env:FAKE_EXCEL_SHEETS -split '\|')) {
      if (-not $s) { continue }
      $sheet = [FakeSheet]::new(); $sheet.Name = $s; $sheet.CodeName = "Sheet$i"
      [void]$wb.Worksheets.List.Add($sheet)
      $d = [FakeComponent]::new(); $d.Name = "Sheet$i"; $d.Type = 100
      $d.Properties.Map['Name'] = @{ Value = $s }
      $comps.Push($d)
      $i++
    }
    if ($env:FAKE_EXCEL_PROJECT) { Import-FakeProject $wb ([IO.File]::ReadAllText($env:FAKE_EXCEL_PROJECT) | ConvertFrom-Json) }
    # 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無いときは VBProject で例外になる
    $wb | Add-Member -MemberType ScriptProperty -Name VBProject -Value {
      if ($env:FAKE_EXCEL_NO_VBOM) { throw 'プログラミングによる Visual Basic プロジェクトへのアクセスは信頼性に欠けます' }
      $this.Project
    }
    $this.Last = $wb
    return $wb
  }
}

class FakeExcel {
  [bool]$Visible = $true
  [bool]$DisplayAlerts = $true
  [bool]$ScreenUpdating = $true
  [bool]$EnableEvents = $true
  [int]$AutomationSecurity = 1
  [int]$Hwnd = 0
  [FakeWorkbooks]$Workbooks = [FakeWorkbooks]::new()
  [void] Quit() {
    # 設定が期待どおりだったかを記録する
    $state = [ordered]@{
      visible = $this.Visible; displayAlerts = $this.DisplayAlerts; enableEvents = $this.EnableEvents
      automationSecurity = $this.AutomationSecurity; closed = $this.Workbooks.Last.Closed
    }
    if ($env:FAKE_EXCEL_LOG) { [IO.File]::WriteAllText($env:FAKE_EXCEL_LOG, ($state | ConvertTo-Json)) }
  }
}

# 取り込みのテスト用: 既存のツールの VBA プロジェクトを再現する
function Import-FakeProject($wb, $spec) {
  $comps = $wb.Project.VBComponents
  foreach ($m in @($spec.components)) {
    $c = $comps.List | Where-Object { $_.Name -eq $m.name } | Select-Object -First 1
    if (-not $c) { $c = [FakeComponent]::new(); $c.Name = $m.name; $c.Type = $m.type; $comps.Push($c) }
    foreach ($l in @($m.code)) { $c.CodeModule.Buf.Add([string]$l) }
    $c.CodeModule.CountOfLines = $c.CodeModule.Buf.Count
    if ($m.form) {
      $c.Designer = @{ Controls = [FakeControls]::new(); Font = @{}; Name = $m.name }
      if ($m.form.picture) { $c.Designer.Picture = @{ Handle = 5 } }
      foreach ($p in $m.form.props.PSObject.Properties) { $c.Properties.Map[$p.Name] = @{ Value = $p.Value } }
      $byName = @{}
      foreach ($k in @($m.form.controls)) {
        $h = @{ __type = $k.type; Name = $k.name; Parent = @{ Name = $k.parent }; Font = @{} }
        foreach ($p in $k.props.PSObject.Properties) {
          if ($p.Name -like 'Font.*') { $h.Font[$p.Name.Substring(5)] = $p.Value } else { $h[$p.Name] = $p.Value }
        }
        if ($k.picture) { $h.Picture = @{ Handle = 7 } }
        if ($k.type -eq 'MultiPage') { $h.Pages = [FakePages]::new() }
        $byName[$k.name] = $h
        if ($k.type -eq 'Page') { [void]$byName[$k.parent].Pages.List.Add($h); $byName[$k.parent].Pages.Count++ }
        else { [void]$c.Designer.Controls.List.Add($h) }
      }
    }
  }
  foreach ($r in @($spec.references)) {
    if ($null -eq $r) { continue }
    [void]$wb.Project.References.List.Add(@{ GUID = $r.guid; Major = $r.major; Minor = $r.minor; BuiltIn = $false; Type = $r.type; Name = $r.name; Description = $r.description })
  }
}

function Get-TypeName($o) {
  if ($o -is [hashtable] -and $o.ContainsKey('__type')) { return $o.__type }
  return 'Unknown'
}

function New-FakeExcel {
  if ($env:FAKE_EXCEL_MISSING) { throw 'クラスが登録されていません' }
  return [FakeExcel]::new()
}
