











<%
FUNCTION PrivacyNoticeLink(FCTR)

    PrivacyNoticeLink = " onclick=" & CHR(34) & "PopUpMsg('Privacy Notice','"
    PrivacyNoticeLink = PrivacyNoticeLink & "Tinctoria Cottage understands you care how information about you is used and shared.   We collect no personal information about you when you visit our web site unless you choose to provide this information to us.  We are not in the business of selling the information you provide to others.  The information collected is used solely by Tinctoria Cottage to provide better services to our friends."
    PrivacyNoticeLink = PrivacyNoticeLink & "'," & FCTR & ");return false" & CHR(34) & " "
    PrivacyNoticeLink = "<a href=''" & PrivacyNoticeLink & "><font size=-1>Privacy Notice</font></a>"

END FUNCTION
%>



