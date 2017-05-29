<%@ LANGUAGE="VBSCRIPT" %>

<% 
session.codepage=65001
'--- Module: FORM.ASP
'---
'--- Simple file upload form, consisting of only a single file upload form element
'---
'--- Copyright (c) 2002 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
%>
<H4>SoftArtisans FileUp Upload Simple Sample</H4>
                  <%
'--- This page will submit the form back to itself. 
'--- The response is rendered by the Request Method.
If Request.ServerVariables("REQUEST_METHOD") <> "POST" THEN 
	'---
	'--- Note the special form definition tag: ENCTYPE="multipart/form-data"
	'---
	%>
<FORM ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST">
                    <TABLE WIDTH="100%" cellpadding=3 cellspacing=0>
                      <TR> 
                        <TD ALIGN="RIGHT" VALIGN="TOP" width=100>Enter Filename:</TD>
                        <%
	'---
	'--- Note the use of the TYPE="FILE" specification
	'---
	%>
                        <TD ALIGN="LEFT"> 
                          <INPUT TYPE="FILE" NAME="FILE1" size="20">
                          <BR>
                          <P><I><B>Note:</B> if a button labeled "Browse..." does 
                            not appear, then your browser does not support File 
                            Upload. For Internet Explorer 3.02 users, a free add-on 
                            is available from Microsoft. Please see the SoftArtisans 
                            FileUp documentation for more information.<BR>
                            </I></P>
                        </TD>
                      </TR>
                      <TR> 
                        <TD ALIGN="RIGHT">&nbsp;</TD>
                        <TD ALIGN="LEFT"> 
                          <INPUT TYPE="SUBMIT" NAME="SUB1" VALUE="Upload File">
                        </TD>
                      </TR>
                    </TABLE>
                  </FORM>
                  <%
Else
	'---
	'--- Instantiate FileUp
	'---
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.codepage=65001
	'---
	'--- Set the default path to store uploaded files. 
	'---
%>
<% if upl.IsEmpty Then %>
                  <P>The file that you uploaded was empty. Most likely, you did 
                    not specify a valid filename to your browser or you left the 
                    filename field blank. Please try again. </P>
                  <% ElseIf upl.ContentDisposition <> "form-data" Then %>
                  <P>Your upload did not succeed, most likely because your browser 
                    does not support Upload via this mechanism.</P>
                  <br>
                  For Internet Explorer Users: 
                  <UL>
                    <LI>For Windows 95 or Windows NT 4.0: 
                      <UL>
                        <LI><A HREF="http://www.microsoft.com/ie/">Download</A> 
                          V3.02 or later of Internet Explorer 
                        <LI><A HREF="http://www.microsoft.com/ie/download">Download</A> 
                          the File Upload Add-on 
                        <LI>For further information, See Knowledge Base Article 
                          <A HREF="http://www.microsoft.com/kb/articles/Q165/2/87.htm">Q165287</A> 
                      </UL>
                    <LI>For Windows 3.1, WFW 3.11 (Windows 16-bit), or Windows 
                      NT 3.51: 
                      <UL>
                        <A HREF="http://www.microsoft.com/ie/">Download</A> V3.02A 
                        or later of Internet Explorer for 16-bit Windows 
                      </UL>
                  </UL>
                  For Netscape Users: 
                  <UL>
                    <LI><A HREF="http://home.netscape.com">Download</A> a version 
                      of Netscape Navigator or Communicator of 2.x or later 
                  </UL>
                  For users of other browsers: 
                  <UL>
                    <LI>Your browser must support a standard called RFC 1867. 
                      Please check with your browser vendor for support of this 
                      standard. 
                  </UL>
                  <%Else %>
                  <P>The file was successfully transmitted by the user.</P>
                  <% 
			on error resume next
			'---
			'--- Save the file now. If you want to preserve the original user's filename, use
			'---
			upl.Save

			'---
			'--- OR, if you want set your own name, uncomment one of the below lines
			'---
			'--- upl.SaveAs "upload.tst" '--- this uses .Path property
			'--- upl.SaveAs "d:\someotherdir\myfile.ext"
			'--- upl.SaveAs "\\bowser\calendar\myfile.ext"	'--- NOTE: The anonymous user *must* have network 
															'--- access rights for UNC names to work
			if Err <> 0 Then %>
                  <H1><FONT COLOR="#ff0000">An error occurred when saving the 
                    file on the server.</FONT></H1>
                  Possible causes include: 
                  <UL>
                    <LI>An incorrect filename was specified 
                    <LI>File permissions do not allow writing to the specified 
                      area 
                  </UL>
                  <P>Please check the FileUp documentation for more troubleshooting 
                    information, or send e-mail to <A HREF="mailto:info@softartisans.com">info@softartisans.com</A>.</P>
                  <%	Else 
				Response.Write("Upload saved successfully.")
				'upl.Delete
			End If %>
                  <P>&nbsp;</P>
                  <CENTER>
                    <FONT SIZE="-1"> 
                    <TABLE WIDTH="80%" BORDER="1" CELLSPACING="1" CELLPADDING="3" HEIGHT="206">
                      <TR> 
                        <TD COLSPAN="2"> 
                          <CENTER>
                            <B>Information About The Uploaded File</B> 
                          </CENTER>
                        </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">&nbsp;User's 
                          filename</TD>
                        <TD WIDTH="70%"><%=upl.UserFilename%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Size 
                          in bytes&nbsp;</TD>
                        <TD WIDTH="70%"><%=upl.TotalBytes%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content 
                          Type</TD>
                        <TD WIDTH="70%"><%=upl.ContentType%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content 
                          Disposition</TD>
                        <TD WIDTH="70%"><%=upl.ContentDisposition%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">MIME 
                          Version</TD>
                        <TD WIDTH="70%"><%=upl.MimeVersion%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Content 
                          Transfer Encoding</TD>
                        <TD WIDTH="70%"><%=upl.ContentTransferEncoding%>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="30%" HEIGHT="27" ALIGN="RIGHT" VALIGN="TOP">Version</TD>
                        <TD WIDTH="70%"><%=upl.Version%>&nbsp;</TD>
                      </TR>
                    </TABLE>
                    </font> 
                    <% 
		End If 
End IF
%>