# PickComputerDialog
Display the 'Select Computer' dialog for Network and Active Directory

<img width="453" height="250" alt="image" src="https://github.com/user-attachments/assets/336b3b27-f679-464e-8c48-9c50be3c7aa6" />

I had always wondered if this was some kind of common dialog made available to other apps through the Windows API, turns out it is and there doesn't appear to be any example of how to call it from VBx/tB, just an incomplete replica, so I wrote a short snippet showing how. There's a lot of options for this dialog so we just go with a simple set.

**Requirements**
- WinDevLib v9.2.627 - Released same day as this snippet: earlier versions had a bug that impacted this project.
- twinBASIC - Unfortunately VB6 people are out of luck for an easy solution as oleexp does not cover the required interfaces, join us in the future :)

```vba
        Dim pPicker As IDsObjectPicker
        Dim hr As Long
        hr = CoCreateInstance(CLSID_DsObjectPicker, Nothing, CLSCTX_ALL, IID_IDsObjectPicker, pPicker)
        If FAILED(hr) Then
            Debug.Print "Failed to create CLSID_DsObjectPicker, 0x" & Hex$(hr)
            Exit Sub
        End If
        
        Dim info As DSOP_INIT_INFO
        info.cbSize = LenB(Of DSOP_INIT_INFO)
        info.cDsScopeInfos = 1
        Dim scope As DSOP_SCOPE_INIT_INFO
        scope.cbSize = LenB(Of DSOP_SCOPE_INIT_INFO)
        scope.flType = DSOP_SCOPE_TYPE_WORKGROUP Or DSOP_SCOPE_TYPE_DOWNLEVEL_JOINED_DOMAIN Or DSOP_SCOPE_TYPE_ENTERPRISE_DOMAIN
        scope.flScope = DSOP_SCOPE_FLAG_DEFAULT_FILTER_COMPUTERS
        scope.FilterFlags.flDownlevel = DSOP_DOWNLEVEL_FILTER_COMPUTERS
        info.aDsScopeInfos = VarPtr(scope)
        
        pPicker.Initialize(info)
        
        Dim spData As IDataObject
        pPicker.InvokeDialog(Me.hWnd, spData)
        If (spData Is Nothing) Or (FAILED(Err.LastHresult)) Then
            Debug.Print "InvokeDialog failed, 0x" & Hex$(Err.LastHresult)
            Exit Sub
        End If
        
        Dim fmt As FORMATETC
        fmt.cfFormat = DCast(Of Integer)(RegisterClipboardFormat(CFSTR_DSOP_DS_SELECTION_LIST))
        fmt.dwAspect = DVASPECT_CONTENT
        fmt.tymed = TYMED_HGLOBAL
        fmt.lIndex = -1
        
        Dim med As STGMEDIUM
        hr = spData.GetData(fmt, med)
        If FAILED(hr) Then
            Debug.Print "Failed to retrieve computer name, 0x" & Hex$(hr)
            Exit Sub
        End If
        
        Dim data As LongPtr = GlobalLock(med.data)
        With CType(Of DS_SELECTION_LIST)(data)
            Dim sRes As String
            sRes = LPWSTRtoStr(.aDsSelection(0).pwzName, False)
        End With
        Text1.Text = sRes
        GlobalUnfix(med.data)
        ReleaseStgMedium(med)

```
