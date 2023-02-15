# 设置开始
$EndTime = Get-Date
$EndTimeStr = $EndTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')


# 设置最后几天数据？
$LastDays = 90

$StartTime = $EndTime.AddDays(-$LastDays)
$StartTimeStr = $StartTime.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')


Write-Host "Gather Citrix Virtual App and Desktop Session Details from StartTime:" $StartTimeStr "to EndTime:" $EndTimeStr "before NOW in Last $LastDays Days."


# 设置访问凭据
$Username = "zjtdemo\ctxadmin"
$Password = ConvertTo-SecureString "abc123!" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)


# 设置Citrix DDC URL
$CitrixDDCURL = "http://192.1.21.101"


# 获取当前时间
$CurrentDate = Get-Date -Format "yyyyMMddTHHmmss"


# 将对象写入 Excel 文件
$Excel = New-Object -ComObject Excel.Application

# 创建一个新的工作簿
$Workbook = $Excel.Workbooks.Add()

# 获取第一个工作表并命名为 "Sheet1"
$Worksheet = $Workbook.Worksheets.Item(1)
$Worksheet.Name = "Sheet1"


# 写入表头
$Worksheet.Cells.Item(1, 1) = 'Start Time'
$Worksheet.Cells.Item(1, 2) = 'End Time'
$Worksheet.Cells.Item(1, 3) = 'Session Type'
$Worksheet.Cells.Item(1, 4) = 'User Name'
$Worksheet.Cells.Item(1, 5) = 'Delivery Group Name'
$Worksheet.Cells.Item(1, 6) = 'Client Address'

# 获取会话信息
$QueryURI =  "$CitrixDDCURL" + '/Citrix/Monitor/OData/v4/Data/Sessions?$expand=Machine($expand=DesktopGroup),User,CurrentConnection' 
$AllSessions = @(Invoke-RestMethod -Method Get -Uri $QueryURI  -Credential $Credential).Value |Where-Object {($_.StartDate -ge $StartTimeStr) -and ($_.EndDate -le $EndTimeStr)}

# 显示会话信息
$row = 2
foreach ($Session in $AllSessions)
{
    $Worksheet.Cells.Item($row, 1) = $Session.StartDate
    $Worksheet.Cells.Item($row, 2) = $Session.EndDate
    $Worksheet.Cells.Item($row, 3) = $Session.SessionType
    $Worksheet.Cells.Item($row, 4) = $Session.User.Username
    $Worksheet.Cells.Item($row, 5) = $Session.Machine.DesktopGroup.Name
    $Worksheet.Cells.Item($row, 6) = $Session.CurrentConnection.ClientAddress
    $row++

}
 
 
$Workbook.SaveAs("C:\Citrix\CitrixSessionDetails-$CurrentDate.xlsx")
$Excel.Quit()  
Write-Host "Gathering information is completed...."
