<div align="center">

## UPDATED\*\*\*   Send Email from within VB full source


</div>

### Description

allows you to send email with one line of code from your vb app.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[xmetrix](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/xmetrix.md)
**Level**          |Beginner
**User Rating**    |3.0 (24 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/xmetrix-updated-send-email-from-within-vb-full-source__1-32494/archive/master.zip)





### Source Code

```
visit www.xmetrix.net/xmail to download the control
'xmail.ocx syntax
xmail1.sendmail [recipient],[from],[subject],[body]
'xmail.ocx usage
Public Sub Command1_Click()
xmail.sendmail "xmetrix@xmetrix.net","test@test.tst","This is my subject", "This should appear in the message of the email."
End Sub
```

