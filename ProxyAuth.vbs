Option Explicit
On Error Resume Next

'�v�����r���������I�u�W�F�N�g
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

'�e�������������N��
objWshShell.Run """C:\kiban\Firefox60\firefox.exe"" -new-tab https://jpn01.safelinks.protection.outlook.com/?url=http%3A%2F%2Fwww.google.co.jp%2F&amp;data=02%7C01%7C%7C1a7c81b72275442e630808d6b637f12f%7Ca4dd529424e441028420cb86d0baae1e%7C1%7C1%7C636896748734633999&amp;sdata=wzs3HZw7K%2BmZWx%2BfaiyaJ8cpfP4AJPmBf2L6UYZfFQg%3D&amp;reserved=0"
WScript.Sleep(3500)

'�e�������������A�N�e�B�u��
objWshShell.AppActivate("Mozilla Firefox")
WScript.Sleep(300)

'�h�c����
objWshShell.SendKeys "CP******"
WScript.Sleep(300)

'�s�`�a���́i�o�v���͗��Ɉړ��j
objWshShell.SendKeys "{TAB}"
WScript.Sleep(300)

'�o�v����
objWshShell.SendKeys "******"
WScript.Sleep(300)

'�d�m�s�d�q����
objWshShell.SendKeys "{ENTER}"

Set objWshShell = Nothing
