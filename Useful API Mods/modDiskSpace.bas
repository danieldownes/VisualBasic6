Attribute VB_Name = "modDiskSpace"
Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
    "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
    lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
    lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters _
    As Long) As Long

Function modFreeSpace() As Integer
    Dim SectorsPerCluster&
    Dim BytesPerSector&
    Dim NumberOfFreeClusters&
    Dim TotalNumberOfClusters&

    dummy& = GetDiskFreeSpace("c:\", SectorsPerCluster, _
    BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters)
    
    modFreeSpace = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
End Function

