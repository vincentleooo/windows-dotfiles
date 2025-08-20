# Optimized Script to focus existing File Explorer window or create new instance
# Description: Fast detection and focus of existing explorer.exe windows

# Add necessary Windows API functions (only once)
if (-not ([System.Management.Automation.PSTypeName]'ExplorerManager').Type) {
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        using System.Text;
        
        public class ExplorerManager {
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            
            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
            
            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
            
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
            
            [DllImport("user32.dll")]
            public static extern bool IsWindowVisible(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool IsIconic(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool BringWindowToTop(IntPtr hWnd);
            
            public const int SW_RESTORE = 9;
            public const int SW_SHOW = 5;
        }
"@
}

# Fast method: Use Get-Process with MainWindowHandle (most efficient)
$explorerWindows = @()

try {
    # Get all explorer processes at once
    $explorerProcesses = @(Get-Process -Name "explorer" -ErrorAction Stop | Where-Object { 
        $_.MainWindowHandle -ne [IntPtr]::Zero -and $_.MainWindowHandle -ne 0
    })
    
    if ($explorerProcesses.Count -gt 0) {
        # Pre-allocate StringBuilder for better performance
        $classNameBuffer = New-Object System.Text.StringBuilder 32
        
        foreach ($proc in $explorerProcesses) {
            $handle = $proc.MainWindowHandle
            
            # Quick visibility check
            if (-not [ExplorerManager]::IsWindowVisible($handle)) { continue }
            
            # Fast class name check
            $classNameBuffer.Clear()
            $result = [ExplorerManager]::GetClassName($handle, $classNameBuffer, 32)
            
            if ($result -gt 0) {
                $className = $classNameBuffer.ToString()
                # Only check for the two main File Explorer classes
                if ($className -eq "CabinetWClass" -or $className -eq "ExploreWClass") {
                    $explorerWindows += @{
                        Handle = $handle
                        ProcessId = $proc.Id
                        ClassName = $className
                    }
                    break  # Found one, that's enough
                }
            }
        }
    }
}
catch {
    # Fallback: No explorer processes found
}

# Fallback method only if primary method fails
if ($explorerWindows.Count -eq 0) {
    # Fast enumeration with early exit
    $foundWindow = $null
    
    $enumCallback = {
        param($hwnd, $lParam)
        
        # Quick checks first (fastest operations)
        if (-not [ExplorerManager]::IsWindowVisible($hwnd)) { return $true }
        
        # Get process ID
        $processId = 0
        [ExplorerManager]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null
        
        # Quick process name check using WMI cache
        try {
            $process = Get-Process -Id $processId -ErrorAction Stop
            if ($process.ProcessName -ne "explorer") { return $true }
        }
        catch { return $true }
        
        # Class name check
        $className = New-Object System.Text.StringBuilder 32
        if ([ExplorerManager]::GetClassName($hwnd, $className, 32) -eq 0) { return $true }
        
        $classStr = $className.ToString()
        if ($classStr -eq "CabinetWClass" -or $classStr -eq "ExploreWClass") {
            $script:foundWindow = @{
                Handle = $hwnd
                ProcessId = $processId
                ClassName = $classStr
            }
            return $false  # Stop enumeration
        }
        
        return $true
    }
    
    [ExplorerManager]::EnumWindows($enumCallback, [IntPtr]::Zero) | Out-Null
    
    if ($script:foundWindow) {
        $explorerWindows += $script:foundWindow
    }
}

# Focus window or create new one
if ($explorerWindows.Count -gt 0) {
    $windowHandle = [IntPtr]$explorerWindows[0].Handle
    
    try {
        # Efficient window focusing
        if ([ExplorerManager]::IsIconic($windowHandle)) {
            [ExplorerManager]::ShowWindow($windowHandle, [ExplorerManager]::SW_RESTORE) | Out-Null
        }
        
        # Single call sequence for best performance
        [ExplorerManager]::BringWindowToTop($windowHandle) | Out-Null
        [ExplorerManager]::SetForegroundWindow($windowHandle) | Out-Null
    }
    catch {
        Start-Process "explorer.exe"
    }
}
else {
    Start-Process "explorer.exe"
}