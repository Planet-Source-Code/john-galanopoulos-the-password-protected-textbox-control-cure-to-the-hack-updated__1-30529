<div align="center">

## The password protected textbox control\. Cure to the hack \(updated\)


</div>

### Description

In the past days, we have seen an exploit on how to get the text behind the password protected textbox control. This was sure a security breach. Windows XP don't have the same problem with their password protected textbox. They use their control from comctl32.dll and not user32.dll! Our textbox has the same protection. It is subclassed to intercept the WM_GETTEXT and EM_SETPASSWORDCHAR by transforming them into WM_NULL upon receival. Download the source and check it out. Any comments or suggestions are always welcomed :-) Subclassing source was borrowed by Stephen Kent's excellent article on PSC site : http://www.planet-source-code.com/xq/ASP/txtCodeId.30275/lngWId.1/qx/vb/scripts/ShowCode.htm
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-01-07 14:26:16
**By**             |[John Galanopoulos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-galanopoulos.md)
**Level**          |Intermediate
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[The\_passwo47035172002\.zip](https://github.com/Planet-Source-Code/john-galanopoulos-the-password-protected-textbox-control-cure-to-the-hack-updated__1-30529/archive/master.zip)








