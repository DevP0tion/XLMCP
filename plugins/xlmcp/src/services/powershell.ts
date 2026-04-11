import { PowerShell } from "node-powershell";

let shell: PowerShell | null = null;

export async function getShell(): Promise<PowerShell> {
  if (!shell) {
    shell = new PowerShell({
      executableOptions: {
        "-ExecutionPolicy": "Bypass",
        "-NoProfile": true,
      },
    });
    // 실행 중인 Excel에 연결 시도, 없으면 새로 생성
    await shell.invoke(`
      try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
      } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
      }
      $excel.DisplayAlerts = $false

      function Resolve-Workbook {
        param([string]$Name)
        if ($Name -and $Name -ne "") {
          return $excel.Workbooks.Item($Name)
        }
        if (-not $excel.ActiveWorkbook) {
          throw "열려 있는 워크북이 없습니다."
        }
        return $excel.ActiveWorkbook
      }

      function Resolve-Sheet {
        param($wb, [string]$SheetName)
        if ($SheetName -and $SheetName -ne "") {
          return $wb.Worksheets.Item($SheetName)
        }
        return $wb.ActiveSheet
      }
    `);
  }
  return shell;
}

export async function runPS(script: string): Promise<string> {
  const ps = await getShell();
  const wrapped = `
    try {
      ${script}
    } catch {
      [Console]::Error.WriteLine(($_ | ConvertTo-Json -Compress))
      throw $_
    }
  `;
  try {
    const result = await ps.invoke(wrapped);
    return result.raw ?? "";
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    const cleaned = msg.replace(/\r?\n/g, " ").trim();

    // 구조화된 에러 추출 시도
    let errorMessage = cleaned;
    const jsonStart = cleaned.indexOf("{");
    const jsonEnd = cleaned.lastIndexOf("}");
    if (jsonStart !== -1 && jsonEnd > jsonStart) {
      try {
        const parsed = JSON.parse(cleaned.slice(jsonStart, jsonEnd + 1));
        errorMessage = parsed.Exception?.Message ?? parsed.FullyQualifiedErrorId ?? cleaned;
      } catch {
        // JSON 파싱 실패 시 원본 메시지 사용
      }
    }

    throw new Error(JSON.stringify({
      error: true,
      message: errorMessage,
      type: "PowerShellError",
    }));
  }
}

export async function dispose(): Promise<void> {
  if (shell) {
    try {
      await shell.invoke(`
        if ($excel) {
          [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
      `);
    } catch { /* ignore */ }
    await shell.dispose();
    shell = null;
  }
}
