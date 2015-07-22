Attribute VB_Name = "Office2010-Module"
' REQUIREMENTS: Windows 98 or above, Access 97 and above
'
' Please read the full tutorial here:
' http://www.everythingaccess.com/tutorials.asp?ID=Get-all-IP-Addresses-of-your-machine
'
' Please leave the copyright notices in place.
' Thank you.
'
'Option Compare Database

'A couple of API functions we need in order to query the IP addresses in this machine
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "Iphlpapi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

'The structures returned by the API call GetIpAddrTable...
Type IPINFO
    dwAddr As Long         ' IP address
    dwIndex As Long         ' interface index
    dwMask As Long         ' subnet mask
    dwBCastAddr As Long     ' broadcast address
    dwReasmSize As Long    ' assembly size
    Reserved1 As Integer
    Reserved2 As Integer
End Type

Declare Function Get_User_Name Lib "advapi32.dll" Alias _
              "GetUserNameA" (ByVal lpBuffer As String, _
               nSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        ByRef nSize As Long) As Long

