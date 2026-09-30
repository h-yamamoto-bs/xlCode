# テスト用の偽の Excel（COM）。build.ps1 の New-Object -ComObject Excel.Application の代わりに使う。
# 保存（SaveAs）すると、VBA プロジェクトの中身を JSON で書き出す。

class FakeCodeModule {
  [System.Collections.Generic.List[string]]$Lines = [System.Collections.Generic.List[string]]::new()
  [int]$CountOfLines = 0
  [void] DeleteLines([int]$start, [int]$count) {
    $this.Lines.RemoveRange($start - 1, $count); $this.CountOfLines = $this.Lines.Count
  }
  [void] AddFromString([string]$code) {
    foreach ($l in ($code -split "`r`n")) { $this.Lines.Add($l) }
    $this.CountOfLines = $this.Lines.Count
  }
}

class FakeProperties {
  [hashtable]$Map = @{}
  [hashtable] Item([string]$name) {
    if (-not $this.Map.ContainsKey($name)) { throw "プロパティ $name はありません" }
    return $this.Map[$name]
  }
}

class FakeControls {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
  [hashtable] Add([string]$progId, [string]$name, [bool]$visible) {
    if (-not $progId.StartsWith('Forms.')) { throw "ProgID が不正です: $progId" }
    $c = @{ ProgId = $progId; Name = $name; Font = @{} }
    if ($progId -eq 'Forms.Frame.1') { $c.Controls = [FakeControls]::new() }
    if ($progId -eq 'Forms.MultiPage.1') {
      $c.Pages = [FakePages]::new()
      [void]$c.Pages.Add('Page1'); [void]$c.Pages.Add('Page2')
    }
    [void]$this.List.Add($c)
    return $c
  }
}

class FakePages {
  [System.Collections.ArrayList]$List = [System.Collections.ArrayList]::new()
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

class FakeProject {
  [FakeComponents]$VBComponents = [FakeComponents]::new()
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

  [void] SaveAs([string]$path, [int]$format) {
    $this.SavedAs = $path
    $this.SavedFormat = $format
    $dump = [ordered]@{ format = $format; autoSave = $this.AutoSaveOn; components = @() }
    foreach ($c in $this.Project.VBComponents.List) {
      $props = [ordered]@{}
      foreach ($k in $c.Properties.Map.Keys) { $props[$k] = $c.Properties.Map[$k].Value }
      $dump.components += [ordered]@{
        name = $c.Name; type = $c.Type; code = @($c.CodeModule.Lines); props = $props; designer = $c.Designer
      }
    }
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

function New-FakeExcel {
  if ($env:FAKE_EXCEL_MISSING) { throw 'クラスが登録されていません' }
  return [FakeExcel]::new()
}
