function Set-WindowState {
	<#
		.SYNOPSIS
			Sets the state of a Window.
	
		.DESCRIPTION
            Set the state such as hidden, maximzed, etc. of a Window.
            WARNING! Don't lose the returned process object if using Hide 
	
		.PARAMETER  WindowCaption
			The caption (name) of a Window to be manipulated.
	
		.PARAMETER  ProcessObject
			The process object having a Window to be manipulated.
	
		.PARAMETER  WindowAction
			The new desired state of the Window.
	
		.EXAMPLE
			Set-WindowState -WindowCaption "Untitled - Notepad" -WindowAction "Minimize"
	
		.EXAMPLE
			Set-WindowState -WindowCaption "Untitled - Notepad" -WindowAction "Normalize"
	
		.EXAMPLE
			$Handle = Set-WindowState -ProcessObject (Get-Process Notepad) -WindowAction "Hide"
	
		.EXAMPLE
			Set-WindowState -ProcessObject $Handle -WindowAction "Normalize"
	
		.INPUTS
			System.String,System.ComponentModel.Process
	
		.OUTPUTS
			System.Array, System.ComponentModel.Process
	
		.NOTES
			Created 2015-05-21, Dennis Lindqvist
	
		.LINK
			https://gallery.technet.microsoft.com/scriptcenter/Demo-of-calling-C-and-6ef0cd2b
	
		.LINK
			https://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx
	
	#>
	
	param (
		[string]$WindowCaption,
		$ProcessObject,
		[ValidateSet("hide","normalize","minimize","maximize","restore","recent","show","show_na","minimize_na","minimize_nx","default","force_min")]
		[string]$WindowAction = "normalize"
	)
	
	switch ($WindowAction)
		{
		"hide"		{$action = 0} 	# $SW_HIDE = 0				Hides the window and activates another window.
		"normalize"	{$action = 1} 	# $SW_SHOWNORMAL = 1		Activates and displays a window
		"minimize" 	{$action = 2} 	# $SW_SHOWMINIMIZED = 2		Activates the window and displays it as a minimized window.
		"maximize" 	{$action = 3} 	# $SW_SHOWMAXIMIZED = 3		Activates the window and displays it as a maximized window.
		"recent" 	{$action = 4} 	# $SW_SHOWNOACTIVATE = 4	Displays a window in its most recent size and position.
		"show" 		{$action = 5}	# SW_SHOW = 5				Activates the window and displays it in its current size and position. 
		"minimize_nx" {$action = 6}	# SW_MINIMIZE = 6			Minimizes the specified window and activates the next top-level window in the Z order
		"minimize_na" {$action = 7}	# SW_SHOWMINNOACTIVE = 7	Displays the window as a minimized window.
		"show_na"	{$action = 8}	# SW_SHOWNA = 8				Displays the window in its current size and position.
		"restore"	{$action = 9}	# SW_RESTORE = 9			Activates and displays the window to its original size and position. 
		"default"	{$action = 10}	# $SW_SHOWDEFAULT = 10		Sets the show state based on the SW_ value specified in the STARTUPINFO structure.
		"force_min"	{$action = 11}	# $SW_FORCEMINIMIZE = 11	Minimizes a window, even if the thread that owns the window is not responding.
	}

	# C# signature for ShowWindowAsync
	$csharpsign = @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@
	
	# Add static methods to this PowerShell session
	$showWindowAsync = Add-Type -MemberDefinition $csharpsign -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

	if (!$ProcessObject) {
		$ProcessObject = (get-process | where {$_.mainwindowtitle -eq $WindowCaption -and $_.Name -ne 'dwm'})
	}
	
	foreach ($Object in $ProcessObject) {
		$showWindowAsync::ShowWindowAsync($Object.MainWindowHandle, $action)
	}
	
	return $ProcessObject # to make normalize possible

# end function Set-WindowState
}