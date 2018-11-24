# Wake On LAN ipv6 (multicast) ipv4 powershell script

# Usage: 
# send-wakeonlan 00:30:1b:42:ea:14 82.109.155.109 8900
# send-wakeonlan 00-0f-0c-34-33-12 192.168.1.255 8900
# send-wakeonlan 00-0f-0c-34-33-12 224.0.0.1 8900
# send-wakeonlan 00-0f-0c-34-33-12 FF02::1 8900
# send-wakeonlan 00-07-Ef-3C-37-22 2001:6f8:366:0:01:f0ff:fe2d:1f24 8900
# send-wakeonlan 00:07:Ef:3C:37:22 2001:6f8:366:056:212:f0ff:fe2d:124 8900
# send-wakeonlan 00:07:Ef:3C:37:22 machine.wakeonlan.fr 8900
# send-wakeonlan 00:07:Ef:3C:37:22 fe80::20C:29ff:fe62:8305%4 8900

param ([String]$macAddress = $(throw 'mac address is required'), [string]$hostname, [int]$port= 8900)

function wakeonlan([string]$macAddress = $(throw 'mac address is required'), [string]$hostname, [int]$port= 8900)
{
  
  if ([System.Net.Sockets.Socket]::OSSupportsIPv6)
  {
    " IPv6 support enabled"
  }
  else
  {
    Write-Warning " Error! IPv6 support not enabled `a"
  }

  $he = [System.Net.Dns]::GetHostEntry($hostname);
  $destAddress= $he.AddressList[0]

  $destination = [System.Net.IPAddress]::Parse($destAddress)

  $endpoint = new-object System.Net.IPEndpoint($destination,$port)

  $socket = new-object System.Net.Sockets.Socket($endpoint.AddressFamily, [System.Net.Sockets.SocketType]::Dgram, [System.Net.Sockets.ProtocolType]::Udp)

  [byte[]]$buffer = [byte[]](,0xFF * 6)

  $buffer += (($macAddress.split('-:') | foreach {[byte]('0x' + $_)}) * 16)

  $sent = $socket.Sendto($buffer, $buffer.length, 0, $endpoint)
  $sent = $socket.Sendto($buffer, $buffer.length, 0, $endpoint)
  $sent = $socket.Sendto($buffer, $buffer.length, 0, $endpoint)

  if ($sent -ne 102)
  {
    Write-Warning " Send error ! `a"
  }

  if ($hostname -eq $destAddress)
  {
    " $sent bytes sent to $hostname port $port.`n `r "
  }
  else
  {
    " $sent bytes sent to $hostname ($destAddress) port $port. `n `r "
  }
  $socket.close()
  $socket= $null
}



send-wakeonlan $macAddress $hostname $port
