//go:build windows

package main

import (
	_ "embed"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
	"unsafe"
)

//go:embed SPED.xlam
var addin []byte

const (
	appName = "SPED"
	version = "1.2.0"
	regKey  = `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\SPED`
)

var (
	user32     = syscall.NewLazyDLL("user32.dll")
	msgBoxProc = user32.NewProc("MessageBoxW")
)

const (
	mbOK         = 0x0
	mbOKCancel   = 0x1
	mbYesNo      = 0x4
	mbIconInfo   = 0x40
	mbIconQ      = 0x20
	mbIconWarn   = 0x30
	mbIconErr    = 0x10
	idOK         = 1
	idYes        = 6
)

func box(text, title string, flags uintptr) int {
	t, _ := syscall.UTF16PtrFromString(text)
	c, _ := syscall.UTF16PtrFromString(title)
	r, _, _ := msgBoxProc.Call(0, uintptr(unsafe.Pointer(t)), uintptr(unsafe.Pointer(c)), flags)
	return int(r)
}

func run(name string, args ...string) (string, error) {
	cmd := exec.Command(name, args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	out, err := cmd.CombinedOutput()
	return strings.TrimSpace(string(out)), err
}

func excelRunning() bool {
	out, _ := run("tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH")
	return strings.Contains(strings.ToUpper(out), "EXCEL.EXE")
}

func waitForExcelClosed(title string) bool {
	for excelRunning() {
		if box("Excel is open.\n\nPlease save your work and close Excel completely, then click OK to continue.", title, mbOKCancel|mbIconWarn) != idOK {
			return false
		}
	}
	return true
}

func addinPath() string {
	return filepath.Join(os.Getenv("APPDATA"), "Microsoft", "AddIns", "SPED.xlam")
}

// PowerShell via Excel COM: turn the add-in on (or off) in Excel's own add-in list.
func setInstalled(on bool) error {
	script := `
$ErrorActionPreference = 'Stop'
$path = $env:SPED_PATH
$on = ($env:SPED_ON -eq '1')
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  $found = $null
  foreach ($a in $xl.AddIns) { if ($a.Name -ieq 'SPED.xlam') { $found = $a } }
  if ($on) {
    if (-not $found) { $found = $xl.AddIns.Add($path, $true) }
    $found.Installed = $true
  } elseif ($found) {
    $found.Installed = $false
  }
  $wb.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
}
`
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	on1 := "0"
	if on {
		on1 = "1"
	}
	cmd.Env = append(os.Environ(), "SPED_PATH="+addinPath(), "SPED_ON="+on1)
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("%v: %s", err, strings.TrimSpace(string(out)))
	}
	return nil
}

func install() {
	title := appName + " Setup"
	if box("Install SPED "+version+" for Microsoft Excel?\n\nThe add-in will be copied to your Excel add-ins folder and switched on. A SPED tab will appear in Excel.", title, mbYesNo|mbIconQ) != idYes {
		return
	}
	if !waitForExcelClosed(title) {
		return
	}
	dst := addinPath()
	if err := os.MkdirAll(filepath.Dir(dst), 0o755); err != nil {
		box("Could not create the add-ins folder:\n"+err.Error(), title, mbOK|mbIconErr)
		return
	}
	if err := os.WriteFile(dst, addin, 0o644); err != nil {
		box("Could not copy the add-in:\n"+err.Error()+"\n\nIs Excel still running in the background?", title, mbOK|mbIconErr)
		return
	}

	// keep a copy of this installer so Windows "Apps" can uninstall it
	self, _ := os.Executable()
	instDir := filepath.Join(os.Getenv("LOCALAPPDATA"), "SPED")
	_ = os.MkdirAll(instDir, 0o755)
	uninst := filepath.Join(instDir, "SPED-Setup.exe")
	if data, err := os.ReadFile(self); err == nil {
		_ = os.WriteFile(uninst, data, 0o755)
	}
	run("reg", "add", regKey, "/v", "DisplayName", "/d", "SPED for Excel", "/f")
	run("reg", "add", regKey, "/v", "DisplayVersion", "/d", version, "/f")
	run("reg", "add", regKey, "/v", "Publisher", "/d", "SPED", "/f")
	run("reg", "add", regKey, "/v", "UninstallString", "/d", `"`+uninst+`" /uninstall`, "/f")
	run("reg", "add", regKey, "/v", "NoModify", "/t", "REG_DWORD", "/d", "1", "/f")

	if err := setInstalled(true); err != nil {
		box("SPED was copied to:\n"+dst+"\n\nbut could not be switched on automatically.\n\nIn Excel go to File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, choose SPED.xlam and tick it.\n\nDetails: "+err.Error(), title, mbOK|mbIconWarn)
		return
	}
	box("SPED "+version+" is installed.\n\nOpen Excel and look for the SPED tab on the ribbon.", title, mbOK|mbIconInfo)
}

func uninstall() {
	title := appName + " Uninstall"
	if box("Remove SPED from Excel?\n\nYour saved cell formats, settings and run log are kept.", title, mbYesNo|mbIconQ) != idYes {
		return
	}
	if !waitForExcelClosed(title) {
		return
	}
	if _, err := os.Stat(addinPath()); err == nil {
		_ = setInstalled(false)
	}
	_ = os.Remove(addinPath())
	run("reg", "delete", regKey, "/f")
	// the running copy cannot delete itself; schedule removal of the folder
	instDir := filepath.Join(os.Getenv("LOCALAPPDATA"), "SPED")
	cmd := exec.Command("cmd", "/C", "ping 127.0.0.1 -n 3 >nul & rmdir /S /Q \""+instDir+"\"")
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	_ = cmd.Start()
	box("SPED has been removed.", title, mbOK|mbIconInfo)
}

func main() {
	for _, a := range os.Args[1:] {
		if strings.EqualFold(a, "/uninstall") || strings.EqualFold(a, "-uninstall") {
			uninstall()
			return
		}
	}
	install()
}
