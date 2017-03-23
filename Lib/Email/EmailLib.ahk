/* 
-----------------------------
Modules' Library
version 0.1
 
Updated: Nov. 04, 2013 
Created by: Feng Wu 
--------------------------------
*/

SetWorkingDir, %A_ScriptDir%

/* 
------------------------------------Free Email----------------------------------
	CreateEmailAccount,
	GetEmailIDFromFreeEmail, 	GetEmailIDByTopic, 	 GetEmailContentFromFreeEmail, 			
	DeleteAccount, 				GetEmailBody,		 GetEmailAttachmentName	
	VerifyEmailKeyInfo,			VerifyEmailAttachmentName,		EmailPageURL
	GoToPageLinkedInEmail,		DeleteEmail,

-------------------------------------------------------------------------------
*/
;-----------------------------
;
; Function: CreateEmailAccount
; Description:
;		Sign up a new email account in "mail.qa.test.com" site
; Syntax: CreateEmailAccount(username, domain [, pwd = "pass"])
; Parameters:
;		username - prefix name of the new email
;		domain - email domain
;		pwd - (Optional)password
; Remarks:
; 		#Include  lib/HTTPRequest.ahk		
; Related:
;		None
; Example: 
;	CreateEmailAccount("feng", "qa.test.com")
; 
;-------------------------------------------------------------------------------
CreateEmailAccount(username, domain, pwd = "pass")
{
   url := "http://mail.qa.test.com/create.asp"

   data =
   ( LTrim Join&
      username=%username%
      domain=%domain%
      pw=%pwd%
      pw1=%pwd%
      reg_Text=
   )
   
  headers =
   ( LTRIM 
		
        Content-Type: application/x-www-form-urlencoded
   )
   
  options =
   ( LTRIM C
      Method: POST
      Charset: utf-8
   
   )
   
   HTTPRequest(url)
   HTTPRequest(url, data, headers, options)
   Debug("[" . A_ThisFunc . "]" . data)
   if data contains User already exists
   {
      
     Debug("[" . A_ThisFunc . "] Err- Maybe [" . username . "@" . domain . "] already exists.")
   }
   else if data contains creation successful
   {
     
     Debug("[" . A_ThisFunc . "] Pass - Create [" . username . "@" . domain . "] successful.")
   }
   else
   {
	  Debug("[" . A_ThisFunc . "] Err- Unknow issue")
   }
 
}

;Change 20120830110849 to 2012-08-30T11:08:49
__GetBeginTime()
{
	FormatTime, beginTimeStr, %A_Now%, yyyy-MM-ddThh:mm:ss
	return beginTimeStr
}

;Add 1 min to now 20120830110849 to 2012-08-30T11:09:49
__GetEndTime() 
{
	mt := A_Min + 1
	mt := mt < 10 ? "0" . mt : mt >= 60 ? "00" : mt
	ht := mt == 00 ? A_Hour + 1 : A_Hour 
	timestr = %A_YYYY%%A_MM%%A_DD%%ht%%mt%%A_Sec%
	FormatTime, endTimeStr, %timestr%, yyyy-MM-ddThh:mm:ss
	return endTimeStr
}

;-----------------------------
;
; Function: GetEmailIDFromFreeEmail
; Description:
;		To get free email ID
; Syntax: GetEmailIDFromFreeEmail(receuser_email)
; Parameters:
;		receuser_email - any valid email address
; Return Value:
;		the latest email id
; Remarks:
;		HELP: [url]http://bk.qa.test.com/freemail/resthelp.txt[/url]
; 		GET http://<url>/freemail/rest/<mailaccount>?offset=&count=&begintime=&endtime=&topiclike=
;		eg: http://bk.qa.test.com/freemail/rest/sp102@sprt66.com?offset=1&count=10&begintime=1345706067000&endtime=1345706097001&topiclike=comment
;		All following parameters are optional.
;       offset: the from index of result range, default 1
;       count:  get count number of mails (MAX 100), default 10
;		begintime: get mails which receive date are >= this date, support format: long, e.g 1345706097001 or "yyyy-MM-dd'T'HH:mm:ss"
;		endtime:   get mails which receive date are <= this date, support format: long, e.g 1345706097001 or "yyyy-MM-dd'T'HH:mm:ss"
;		topiclike:   get mails which topic contains this string.
; Related:
;		GetEmailContentFromFreeEmail
; Example: 
;		[code]emailID := GetEmailIDFromFreeEmail("anyemail@test.com")[/code]
; 
;-------------------------------------------------------------------------------
GetEmailIDFromFreeEmail(receuser_email)
{
   global @error_msg
   url := "http://bk.qa.test.com/freemail/rest/" . receuser_email
   ; begintime = __GetBeginTime()
   ; endtime = __GetEndTime()
   data =
   ( LTrim Join&
      offset = 1
      count = 1
   )
   
  options =
   ( LTRIM C
      Method: GET
      Charset: utf-8   
   )
   
   HTTPRequest(url, jdata := data,, options)
   if (jdata != "[]")
   {
	  return json(jdata, "id")
   }
   else
   {
	 @error_msg := "[" . A_ThisFunc . "] Err - No emails were found in " .  receuser_email . " account `r`n"
	 Gosub, ErrorLabel  
   }
   
}

;-----------------------------
;
; Function: GetEmailContentFromFreeEmail
; Description:
;		To get email content
; Syntax: GetEmailContentFromFreeEmail(email_id)
; Parameters:
;		email_id 
; Return Value:
;		json string
; Remarks:
;		http://<url>/freemail/rest/mails/<mailid>  
;		eg:http://bk.qa.test.com/freemail/rest/mails/sp102%40sprt66.com187533108383184
; Related:
;		GetEmailInfoFromFreeEmail
; Example: 
;		[code] jdata := GetEmailInfoFromFreeEmail(email)
;		eid := json(jdata, "id")
;		email_content := GetEmailContentFromFreeEmail(eid)[/code]
; 
;-------------------------------------------------------------------------------
GetEmailContentFromFreeEmail(email_id)
{
   encid := UriEncode(email_id)
   url := "http://bk.qa.test.com/freemail/rest/mails/" . encid
   HTTPRequest(url, email_content)
   Debug("[" . A_ThisFunc . "] Pass - emailID=" . email_id)
   return email_content
}

;-----------------------------
;
; Function: DeleteEmail
; Description:
;		Delete an email 
; Syntax: DeleteEmail(email_id)
; Parameters:
;		email_id 
; Return Value:
;		None
; Remarks:
;		delete an email (be careful, cannot recover) 
;       DELETE http://<url>/freemail/rest/mails/<mailid>  
; Related:
;		GetEmailInfoFromFreeEmail
; Example: 
;		[code] jdata := GetEmailInfoFromFreeEmail(email)
;		eid := json(jdata, "id")
;		DeleteEmail(eid)[/code]
; 
;-------------------------------------------------------------------------------
DeleteEmail(email_id) 
{
   options =
   ( LTRIM C
      Method: DELETE
   )
   url := "http://bk.qa.test.com/freemail/rest/mails/" . email_id
   HTTPRequest(url,data,, options)
   Debug("[" . A_ThisFunc . "]" . (data == 1 ? "Pass!!" : "Err!"))
}

;-----------------------------
;
; Function: DeleteAccount
; Description:
;		Delete an email account and delete all emails 
; Syntax: DeleteAccoun(email)
; Parameters:
;		email_id 
; Return Value:
;		None
; Remarks:
;		delete ALL emails of an account (be careful!!! cannot recover)
;       DELETE http://<url>/freemail/rest/<mailaccount>
; Related:
;		DeleteEmail
; Example: 
;		[code]DeleteAccount("anyemail@test.com")[/code]
; 
;-------------------------------------------------------------------------------
DeleteAccount(email)
{
   options =
   ( LTRIM C
      Method: DELETE
   )
   url := "http://bk.qa.test.com/freemail/rest/" . email
   HTTPRequest(url,data,, options)
   Debug("[" . A_ThisFunc . "] " . (data == 1 ? "Pass!" : "Err!"))
}

;-----------------------------
;
; Function: VerifyEmailKeyInfo
; Description:
;		verify email key information such as email subject, sender, senderName, reply-to
; Syntax: VerifyEmailKeyInfo(key, template_value, email_content)
; Parameters:
;		key - subject, sender, senderName
;       email_content - json data in free email server response
;		template_value - value in the email template
; Return Value:
;		None
; Remarks:
;		Don't block TA case execution if failed 
; Related:
;		GetEmailInfoFromFreeEmail, GetEmailContentFromFreeEmail, GetEmailBody
; Example: 
;		[code] jdata := GetEmailInfoFromFreeEmail(email,0,0)
;		eid := json(jdata, "id")
;		email_content := GetEmailContentFromFreeEmail(eid)
;		VerifyEmailKeyInfo("subject", "Action required: Reset your account password", email_content)[/code]
;
;-------------------------------------------------------------------------------
VerifyEmailKeyInfo(key, email_content, template_value)
{
	global @error_msg
    content_value := json(email_content, key)
   if (template_value == content_value)
   {
      Debug("[" . A_ThisFunc . "] Pass - " . key )
   }
   else
   {
      @error_msg := "[" . A_ThisFunc . "] Err - " . key . " != [" . template_value . "]"
      Gosub, WarringLabel
   }
}

;-----------------------------
;
; Function: GetEmailBody
; Description:
;		verify email key information such as email subject, sender, senderName, reply-to
; Syntax: GetEmailBody(email_content [, format])
; Parameters:
;		email_content - json data
;       format - "HTML" html email "TEXT" plain text email
; Return Value:
;		email body in free email server response
; Remarks:
;		None
; Related:
;		GetEmailInfoFromFreeEmail, GetEmailContentFromFreeEmail, VerifyEmailKeyInfo
; Example: 
;		[code] jdata := GetEmailInfoFromFreeEmail(email,0,0)
;		eid := json(jdata, "id")
;		email_content := GetEmailContentFromFreeEmail(eid)
;		GetEmailBody(email_content, "HTML")[/code]
;
;-------------------------------------------------------------------------------
GetEmailBody(email_content, format = "HTML")
{
   StringUpper, upperstr, format
   if(upperstr == "HTML"){
		emailBody := json(email_content, "parts[0].content") 
		StringReplace, emailBody, emailBody, `\n , , All 	
		StringReplace, emailBody, emailBody, `\t , , All 	
		StringReplace, emailBody, emailBody, `\" , `",All 	
   }else{
		emailBody := json(email_content, "parts[1].content")
		StringReplace, emailBody, emailBody, `\n , , All 	
   }
   
   return RegExReplace(emailBody, "(?<=\>)(?:\s*\r?\n?)(?=\<)", "")		; Remove all whitespace characters between  html tags
   ;return emailBody
   
}

;-----------------------------
;
; Function: GetEmailAttachmentName
; Description:
;		Get email attachment name
; Syntax: GetEmailAttachmentName(email_content)
; Parameters:
;		email_content - json data
; Return Value:
;		email attachment name
; Remarks:
;		None
; Related:
;		VerifyEmailAttachmentName
; Example: 
;		jdata := GetEmailInfoFromFreeEmail(email,0,0)
;		eid := json(jdata, "id")
;		email_content := GetEmailContentFromFreeEmail(eid)
;		attachment_name = GetEmailAttachmentName(email_content)
;
;-------------------------------------------------------------------------------
GetEmailAttachmentName(email_content){
	return json(email_content, "attachments[0].name")
}

;-----------------------------
;
; Function: VerifyEmailAttachmentName
; Description:
;		verify email key information such as email subject, sender, senderName, reply-to
; Syntax: VerifyEmailAttachmentName(template_value, email_content)
; Parameters:
;       email_content - json data in free email server response
;		template_value - value in the email template
; Return Value:
;		None
; Remarks:
;		Don't block TA case execution if failed 
; Related:
;		GetEmailAttachmentName
; Example: 
;		jdata := GetEmailInfoFromFreeEmail(email,0,0)
;		eid := json(jdata, "id")
;		email_content := GetEmailContentFromFreeEmail(eid)
;		VerifyEmailAttachmentName("TA_Meeting Topic.ics", email_content)
;
;-------------------------------------------------------------------------------
VerifyEmailAttachmentName(email_content, template_value)
{
   global @error_msg
   attachment_name := GetEmailAttachmentName(email_content)
   if (template_value == attachment_name)
   {
      Debug("[" . A_ThisFunc . "] Pass - " . attachment_name )
   }
   else
   {
	  @error_msg := "[" . A_ThisFunc  . "] Err - [ " . template_value . " != " . attachment_name . "]"
      Gosub, WarringLabel
   }
}


; if there are multiple emails to be matched and with the same subject, 
; so we should select the one according to email_code in email body,
__GetEmailIDByEmailCode(email, email_code)
{
	eid := json(email, "id")
	if (!!email_code)
    {
		email_content := GetEmailContentFromFreeEmail(eid)
		plain_text := GetEmailBody(email_content, "Text")
		StringCaseSense, On
		IfInString, plain_text, %email_code%
		{
			return eid			; found the email
		}
		else
		{
			return false	
		}
    }
	else
	{
		return eid
	}

}

; To find email id according to email code
__SearchEmailByEmailCode(jarr, email_code)
{
	global @error_msg
   For index, email in jarr
   {
      r := __GetEmailIDByEmailCode(email, email_code)
      if (!!r)                   
      {
         break										; if matched email_code, return email_ID immediately
      }
   }
      
   if (!!r)
   {
      return r
   }
   else
   {
      ;Debug("[" . A_ThisFunc . "] Err - No matched EmailCode [" . email_code . "]")
	  @error_msg := "[" . A_ThisFunc . "] Err - No matched EmailCode [" . email_code . "] `r`n"
	  Gosub, ErrorLabel  
   }
   
}

;-----------------------------
;
; Function: GetEmailIDByTopic (v1.2.1)
; Description:
;		Get email id according to email subject and email code
; Syntax: GetEmailIDByTopic(receuser_email, key_str[, email_code])
; Parameters:
;       receuser_email - any valid email address
;		key_str - emails which subjcet contain the string (case sensitive)
;		email code - To distinguish email easier, hope to add the email code to the subject, eg: MT-A-001
; Return Value:
;		Matched email id
; Remarks:
;		Don't add "{}" in the email subject  
; Related:
;		GetEmailIDFromFreeEmail
; Example: 
;		eid := GetEmailIDByTopic("sp102@sprt66.com", "Invitation", "MT-A-001")
;
;-------------------------------------------------------------------------------
GetEmailIDByTopic(receuser_email, key_str, email_code = "")
{
   global @error_msg
   ;url := "http://bk.qa.test.com/freemail/rest/" . receuser_email
   url := "http://bk.qa.test.com/freemail/rest/" . receuser_email . "?topiclike=" . key_str
  /* 
   data =
   ( LTrim Join&
      topiclike=%key_str%
   )
   */
  options =
   ( LTRIM C
      Method: GET
      Charset: utf-8   
   )
   
   ;HTTPRequest(url, json_str := data, headers := "", options)
   HTTPRequest(url, json_str := "", headers := "", options)
   ;HTTPRequest(url, json_str := data,, options)
   if (json_str != "[]")
   {
     ; StringReplace, jdata, jdata, `\" , `",All 	
	  jarr := ParseJsonStrToArr(json_str) 
      return __SearchEmailByEmailCode(jarr, email_code)
   }
   else
   {
		@error_msg := "[" . A_ThisFunc . "] Err - No emails were found in " .  receuser_email . " account `r`n"
		Gosub, ErrorLabel     
   }
   
}

;-----------------------------
;
; Function: EmailPageURL(v1.2.1)
; Description:
;		Show the email detail page
; Syntax:  EmailPageURL(email_id)
; Parameters:
;       email_id 
; Return Value:
;		None
; Remarks:
;		Default page is html format   
; Related:
;		GetEmailIDFromFreeEmail, GetEmailIDByTopic
; Example: 
;		eid := GetEmailIDByTopic("sp102@sprt66.com", "Invitation", "MT-A-001")
;		email_page_url := EmailPageURL(eid)
;
;-------------------------------------------------------------------------------
EmailPageURL(email_id)
{
   return "http://bk.qa.test.com/freemail/?cmd=view&messid=" . email_id
   
}

;-----------------------------
;
; Function: GoToPageLinkedInEmail (v1.2.1)
; Description:
;		Show the target page linked in email detail page
; Syntax:  GoToPageLinkedInEmail(email_body, begin_str, end_str)
; Parameters:
;       email_body - email html code returned by server 
;		begin_str - String beginning at email body
;		end_str	 -  String ending in email body
; Return Value:
;		None
; Remarks:
;		Don't get the email html code from firebug, because blank characters affect matched results  
; Related:
;		GetEmailIDFromFreeEmail, GetEmailIDByTopic, GetEmailContentFromFreeEmail,  GetEmailBody
; Example: 
;		 eid :=GetEmailIDByTopic("sp102@sprt66.com", "Invitation", "22")
;		 email_content := GetEmailContentFromFreeEmail(eid)
;		 email_body := GetEmailBody(email_content, "HTML")
;		 begin_str =
;		(`
;			<div style="border:1px;font-family: Arial;width:378px; overflow:hidden"><a href="
;		)
;		end_str=
;		(`
;		   " style="font-family
;		 )
;		GoToPageLinkedInEmail(email_body, begin_str, end_str)
;
;-------------------------------------------------------------------------------
GoToPageLinkedInEmail(email_body, begin_str, end_str)
{
   global @error_msg 
   url := GetValFromStr(email_body, begin_str, end_str)
   if(!!url)
   {
      OpenPageWithBrowser(url)
   }
   else
   {
      @error_msg := "[" . A_ThisFunc . "] Err - Gotten link is blank, please check 'beginStr' and 'endStr' again . `r`n"
      Gosub, ErrorLabel
   }
   
}



/* 
------------------------------------Outlook Email----------------------------------
GetFirstEmailFromOL, 			SearchEmailByTopicFromOL, 	InstantSearchEmailFromOL
SearchEmailByTopicFromOL_AD, 	DeleteEmailFromOL, 			DeleteAllEmailsFromOL
SendEmailFromOL,				ReceiveEmailFromOL
-------------------------------------------------------------------------------
*/

; Get folder object (Inbox) in specified store (Account Settings -> Data Files -> Name in Outlook Data File dialog)
__GetStoreInboxFolderObj(data_file_name)
{
	global @error_msg
	try
	{
		stores := ComObjActive("Outlook.Application").Session.Stores
	}
	catch 
	{
		Debug("[" . A_ThisFunc . "] Err- Please open Outlook first, then try again")
		ExitApp
	}
	
	if (!stores.Count)
	{
		@error_msg := "[" . A_ThisFunc . "] Err - No matched store [" . data_file_name . "], please check the Outlook Data File Name `r`n"
		Gosub, ErrorLabel

	}
	else
	{
		Loop % stores.Count
		{
			if (stores.Item(A_Index).DisplayName == data_file_name)
			{
				store := stores.Item(A_Index)
				; only available to outlook2010
				; olFolderInbox = 6
				; store_Inbox_obj := store.GetDefaultFolder(olFolderInbox)  
				folders := store.GetRootFolder.Folders
				Loop % folders.Count
				{
					if (folders.Item(A_Index).Name == "Inbox")
					{
						obj_inbox_folder := folders.Item(A_Index)
						break
					}
				
				}
				break
			}
			
		}
		
		if (obj_inbox_folder)
		{
			return obj_inbox_folder
		}
		else
		{
			@error_msg := "[" . A_ThisFunc . "] Err -  Store [" . data_file_name . "] has no Inbox folder `r`n"
			Gosub, ErrorLabel
		}
		
	}
  
}

;-----------------------------
;
; Function: GetFirstEmailFromOL (v1.2.2)
; Description:
;		Get the first email from inbox of the store for testing
; Syntax:  GetFirstEmailFromOL(data_file_name)
; Parameters:
;       data_file_name - (Account Settings -> Data Files -> Name in Outlook Data File dialog)
; Return Value:
;		MailItem object
; Remarks:
;		MailItem object outlook2007 : http://msdn.microsoft.com/en-us/library/office/bb176688%28v=office.12%29.aspx
;						outlook2010: http://msdn.microsoft.com/en-us/library/ff861252.aspx
; Related:
;		InstantSearchEmailFromOL, SearchEmailByTopicFromOL, SearchEmailByTopicFromOL_AD
; Example: 
;		email_item := GetFirstEmailFromOL("feng@hfqa11ef11.com")
;		email_item.BodyFormat := 2			; olFormatHTML
;		Debug(email_item.HTMLBody)
;		email_item.BodyFormat := 1			; olFormatPlain
;		Debug(email_item.Body)
;
;-------------------------------------------------------------------------------
GetFirstEmailFromOL(data_file_name)
{
	global @error_msg
	olMail = 43
	obj_mail_items := __GetStoreInboxFolderObj(data_file_name).Items
	obj_mail_item := obj_mail_items.GetFirst()
	/*
	While (obj_mail_item.Class != olMail and obj_mail_item != "")
	{
		item := obj_mail_item.GetNext()
	}
	*/
	try
	{
		obj_mail_item.Display(False)
		;WinWaitActive, ahk_class rctrl_renwnd32,,1  active email message window
		return obj_mail_item
	}catch {
		@error_msg := "[" . A_ThisFunc . "] Err - Inbox folder is empty `r`n"
		Gosub, ErrorLabel
	}

}

;-----------------------------
;
; Function: InstantSearchEmailFromOL (v1.2.2)
; Description:
;		only search an email according to keyword from inbox of the store for testing
; Syntax:  InstantSearchEmailFromOL(data_file_name, keyword)
; Parameters:
;       data_file_name - (Account Settings -> Data Files -> Name in Outlook Data File dialog)
;		keyword - A search string that can contain any valid keywords supported in Instant Search.
; Return Value:
;		MailItem object
; Remarks:
;		valid keyword refer to http://technet.microsoft.com/en-us/library/cc513841.aspx
;		the method is limited to search in subject 
; Related:
;		GetFirstEmailFromOL, SearchEmailByTopicFromOL, SearchEmailByTopicFromOL_AD
; Example: 
;		email_item := InstantSearchEmailFromOL("feng@hfqa11ef11.com", "crdc")
;		email_item.BodyFormat := 2			; olFormatHTML
;		Debug(email_item.HTMLBody)
;		email_item.BodyFormat := 1			; olFormatPlain
;		Debug(email_item.Body)
;
;-------------------------------------------------------------------------------
InstantSearchEmailFromOL(data_file_name, keyword)
{
		global @error_msg
		obj_store_inbox_folder := __GetStoreInboxFolderObj(data_file_name)
		app := obj_store_inbox_folder.Application
		active_explorer := app.ActiveExplorer()
		active_explorer.CurrentFolder := __GetStoreInboxFolderObj(data_file_name)
		olSearchScopeCurrentFolder = 0
		;~~ http://technet.microsoft.com/en-us/library/cc513841.aspx
		active_explorer.Search("subject:" . keyword, olSearchScopeCurrentFolder)
		active_explorer.Display()
		try
		{
			obj_mail_item := active_explorer.Selection[1]
			obj_mail_item.Display(False)
			return obj_mail_item
		}
		catch
		{
			@error_msg := "[" . A_ThisFunc . "] Err - No search results, keyword is [" . keyword . "] `r`n"
			Gosub, ErrorLabel
		}
}

;-----------------------------
;
; Function: SearchEmailByTopicFromOL (v1.2.2)
; Description:
;		only search an email according to keyword from inbox of the store for testing
; Syntax:  SearchEmailByTopicFromOL(data_file_name, keyword)
; Parameters:
;       data_file_name - (Account Settings -> Data Files -> Name in Outlook Data File dialog)
;		keyword - A search string that can contain any valid keywords supported in Instant Search.
; Return Value:
;		MailItem object
; Remarks:
;		valid keyword refer to http://technet.microsoft.com/en-us/library/cc513841.aspx
;		the method is limited to search in subject 
; Related:
;		GetFirstEmailFromOL, InstantSearchEmailFromOL, SearchEmailByTopicFromOL_AD
; Example: 
;		email_item := SearchEmailByTopicFromOL("feng@hfqa11ef11.com", "crdc")
;		email_item.BodyFormat := 2			; olFormatHTML
;		Debug(email_item.HTMLBody)
;		email_item.BodyFormat := 1			; olFormatPlain
;		Debug(email_item.Body)
;
;-------------------------------------------------------------------------------
SearchEmailByTopicFromOL(data_file_name, keyword)
{
    global @error_msg
	obj_store_inbox_folder := __GetStoreInboxFolderObj(data_file_name)
	filter := "@SQL=urn:schemas:mailheader:subject like '%" . keyword . "%'"
	obj_table := obj_store_inbox_folder.GetTable(filter)
	app := obj_store_inbox_folder.Application
	
	if (!obj_table.GetRowCount)
	{
		@error_msg := "[" . A_ThisFunc . "] Err - No matched emails [" . keyword . "]`r`n"
		Gosub, ErrorLabel
	}
	else
	{
		loop % obj_table.GetRowCount
		{
			obj_row := obj_table.GetNextRow()
			;~~ http://msdn.microsoft.com/en-us/library/office/bb176406%28v=office.12%29.aspx
			entry_id := obj_row.Item("EntryID")
			obj_mail_item := app.Session.GetItemFromID(entry_id)
			obj_mail_item.Display(False) 
			return obj_mail_item
		}
	
	}
	
}
;-----------------------------
;
; Function: SearchEmailByTopicFromOL_AD (v1.2.2)
; Description:
;		only search an email according to keyword from inbox of the store for testing 
; Syntax:  SearchEmailByTopicFromOL_AD(data_file_name, keyword)
; Parameters:
;       data_file_name - (Account Settings -> Data Files -> Name in Outlook Data File dialog)
;		keyword - A search string that can contain any valid keywords supported in Instant Search.
; Return Value:
;		MailItem object
; Remarks:
;		valid keyword refer to http://technet.microsoft.com/en-us/library/cc513841.aspx
;		the method is limited to search in subject 
;		Retrieve the results asynchronously, Application.AdvancedSearch method and handle the Application.AdvancedSearchComplete event 
; Related:
;		GetFirstEmailFromOL, InstantSearchEmailFromOL, SearchEmailByTopicFromOL
; Example: 
;		email_item := SearchEmailByTopicFromOL_AD("feng@hfqa11ef11.com", "crdc")
;		email_item.BodyFormat := 2			; olFormatHTML
;		Debug(email_item.HTMLBody)
;		email_item.BodyFormat := 1			; olFormatPlain
;		Debug(email_item.Body)
;
;-------------------------------------------------------------------------------
SearchEmailByTopicFromOL_AD(data_file_name, keyword)
{
		;global error_msg, searched_email
		global @error_msg
		obj_store_inbox_folder := __GetStoreInboxFolderObj(data_file_name)
		app := obj_store_inbox_folder.Application
		filter := "urn:schemas:mailheader:subject like '%" . keyword . "%'"
		;filter := "urn:schemas:httpmail:subject LIKE '%" . keyword . "'"
		;~~ http://schemas.microsoft.com/mapi/proptag/0x0037001E (the Subject property referenced by the MAPI ID namespace) 
		;~~ http://msdn.microsoft.com/en-us/library/office/bb147567%28v=office.12%29.aspx
		;filter :=  "http://schemas.microsoft.com/mapi/proptag/0x0037001E" . " ci_phrasematch 'CRDC'" 
		store_Inbox_path  :=  obj_store_inbox_folder.FolderPath
		scope := "'" . store_Inbox_path .  "'"
		sch := app.AdvancedSearch(scope, filter, False)
		ComObjConnect(app, "_app_")
		Sleep, 500
		rsts :=  sch.Results
		Loop % rsts.Count
		{
			searched_email := rsts.Item(A_Index)
			if (searched_email != "")
			{
				searched_email.Display(False) 
			}else{
				@error_msg := "[" . A_ThisFunc . "] Err - Get email failed, [" . keyword . "] was not found! `r`n"
				Gosub, ErrorLabel
			}
		}	
	ComObjConnect(app)
	return searched_email
}

_app_AdvancedSearchComplete(sch)
{
	Debug("[" . A_ThisFunc . "] Done! [" . sch.Results.Count . "]`r`n")
}
;-----------------------------
;
; Function: DeleteEmailFromOL (v1.2.2)
; Description:
;		Delete an email
; Syntax:  DeleteEmailFromOL(obj_mail_item)
; Parameters:
;       obj_mail_item - MailItem object
; Return Value:
;		None
; Remarks:
;		None
; Related:
;		DeleteAllEmailsFromOL
; Example: 
;		email_item := SearchEmailByTopicFromOL("feng@hfqa11ef11.com", "crdc")
;		DeleteEmailFromOL(email_item)
;
;-------------------------------------------------------------------------------
DeleteEmailFromOL(obj_mail_item)
{
	obj_mail_item.Delete()
	
}

;-----------------------------
;
; Function: DeleteAllEmailsFromOL (v1.2.2)
; Description:
;		Delete all emails from inbox of the store for testing
; Syntax:  DeleteAllEmailsFromOL(data_file_name)
; Parameters:
;       data_file_name - (Account Settings -> Data Files -> Name in Outlook Data File dialog)
; Return Value:
;		None
; Remarks:
;		None
; Related:
;		DeleteEmailFromOL
; Example: 
;		DeleteAllEmailsFromOL("feng@hfqa11ef11.com")
;
;-------------------------------------------------------------------------------
DeleteAllEmailsFromOL(data_file_name)
{
	global @error_msg
	obj_mail_items := __GetStoreInboxFolderObj(data_file_name).Items
	if (!obj_mail_items.Count)
	{
		@error_msg := "[" . A_ThisFunc . "] Warring - There is no emails in Inbox of the [" . data_file_name . "] to be deleted `r`n"
		Gosub, WarringLabel
	}
	
	Loop  % obj_mail_items.Count
	{
		obj_mail_items.Item(A_Index).Delete()
	}
	
}


;-----------------------------
;
; Function: SendEmailFromOL 
; Description:
;		Send an email from Outlook
; Syntax: SendEmailFromOL(to, subject, body [, attachments = "", mailAccountName = "", isShown = false])
; Parameters:
;		to - Recipient(s) email such as "feng@test.com", "feng@test.com;feng2@test.com"
;		subject - Email subject
;		body - Email body
;		attachments - Attached file(s) which delimiters either ";" or ","
;			such as "C:\a.txt", "C:\a.txt;", "C:\a.txt,", "C:\a.txt;
;			C:\b.txt", "C:\a.txt;C:\b.txt;"
;		mailAccountName - Send email by other account in "Account Settings"
;		isShown - false (Default): Send email automatically without any prompt
;			true: Show new message window, and the email is only to be sent by manual	
; Remarks:
; 		Whether the email has been sent successfully or not, Check "Sent Items" directory
; Related:
;		ReceiveEmailFromOL
; Example: 
;		body=
;		(
; 			mutil lines text on the email body
;		)
;		; Sender is default account: feng@test.com
;		SendEmailFromOL("feng@test.com;feng2@test.com", "Subject for test email", body) 
;		; Sender is other account: jon@mail.qa.test.com, Jon is the account name in "Account Settings"
;		SendEmailFromOL("feng@test.com", "Subject for test email", body, "C:\a.txt", "Jon", true)
; 
;-------------------------------------------------------------------------------
SendEmailFromOL(to, subject, body, attachments = "", mailAccountName = "", isShown = false)
{
   /* Check  outlook email client
   Loop, HKEY_CURRENT_USER, Software\Microsoft\Office\11.0, 2   ;<-- replace 11.0 with your Office version
     msgbox %A_LoopRegName%
  exitapp
  
  RegRead, OutLookVer, HKCR, Outlook.Application\CurVer ;http://en.wikipedia.org/wiki/Microsoft_Outlook
 */
  olMailItem := 0
  try
    {
      nsp := ComObjActive("Outlook.Application").GetNamespace("MAPI")  
    }
  catch 
  {
	  Debug("[" . A_ThisFunc . "] Err- Please open Outlook first, then try again")
  }
  
  objMail := ComObjActive("Outlook.Application").CreateItem(olMailItem)
  
  if(!!mailAccountName)															; mailAccountName != ""
  {
    objMail.SendUsingAccount := nsp.Accounts.Item(mailAccountName)          ; change from to other mail account name, which is the name in "Account Settings"
  }
 
  objMail.Subject := subject                                                    ; Subject 
  objMail.To := to                                                              ; Recipient 
  objMail.HTMLBody := body                                                      ; Email message body

  if (!!attachments)                                                        	; Add attachments
  {
    Loop, parse, attachments, `;`,                                              ; Delimiters of attachments
    {
      if (!!A_LoopField)
      {
        try
        {
          objMail.attachments.add(A_LoopField)
        }
        catch e
        {
			Debug("[" . A_ThisFunc . "] Err - " . e.Message)       
        }
      }
    }
  }
  
  if(isShown)                                                                   ; Show new email window
  {
    try
    {
      objMail.Display(true)
    }
    catch e
    {
		Debug("[" . A_ThisFunc . "] Err - " . e.Message)
    }
        
  }
  else
  {
    objMail.Send
  }
  
}

;-----------------------------
;
; Function: ReceiveEmailFromOL (v1.2.2)
; Description:
;		Initiates immediate delivery of all undelivered messages submitted in the current session, and immediate receipt of mail for all accounts in the current profile. 
; Syntax: ReceiveEmailFromOL()
; Parameters:
;		None	
; Remarks:
; 		The method will receive mails for all folders
; Related:
;		SendEmailFromOL
; Example: 
;		ReceiveEmailFromOL()
; 
;-------------------------------------------------------------------------------
ReceiveEmailFromOL()
{
	try
    {
		nsp:= ComObjActive("Outlook.Application").GetNamespace("MAPI")  
		nsp.SendAndReceive(true)
	}
	catch 
	{
	  Debug("[" . A_ThisFunc . "] Err- Please open Outlook first, then try again")
	  Gosub, ErrorLabel
	}
	
	WinWaitActive, ahk_class #32770,,3 	; Outlook Send/Receive Progress window
	WinWaitClose,,,5
}

;-----------------------------
;
; Function: GetHTMLEmailBodyFromOL (v1.2.2)
; Description:
;		Get HTML body of email 
; Syntax: GetHTMLEmailBodyFromOL(obj_mail_item)
; Parameters:
;		obj_mail_item - MailItem Object (http://msdn.microsoft.com/en-us/library/ff870912.aspx)
; Remarks:
; 		Remove all whitespace characters between html tags
; Related:
;		GetPlainTextEmailBodyFromOL
; Example: 
;		e:= GetFirstEmailFromOL("feng@wtv12ef.com")
;		html := GetHTMLEmailBodyFromOL(e)
; 
;-------------------------------------------------------------------------------
GetHTMLEmailBodyFromOL(obj_mail_item)
{
	obj_mail_item.BodyFormat := 2			; olFormatHTML
	html := obj_mail_item.HTMLBody
	return RegExReplace(html, "(?<=\>)(?:\s*\r?\n?)(?=\<)", "") ; Remove all whitespace characters between  html tags
}

;-----------------------------
;
; Function: GetPlainTextEmailBodyFromOL (v1.2.2)
; Description:
;		Get plain text of email 
; Syntax: GetPlainTextEmailBodyFromOL(obj_mail_item)
; Parameters:
;		obj_mail_item - MailItem Object
; Remarks:
; 		N/A
; Related:
;		GetHTMLEmailBodyFromOL
; Example: 
;		e:= GetFirstEmailFromOL("feng@wtv12ef.com")
;		text := GetPlainTextEmailBodyFromOL(e)
; 
;-------------------------------------------------------------------------------
GetPlainTextEmailBodyFromOL(obj_mail_item)
{
	obj_mail_item.BodyFormat := 1			; olFormatPlain
	return obj_mail_item.Body
}

;-----------------------------
;
; Function: SwitchMessageWinFromOL (v1.2.2)
; Description:
;		Active an email message window
; Syntax: SwitchMessageWinFromOL()
; Parameters:
;		None
; Remarks:
; 		Email message window should have been opened
; Related:
;		
; Example: 
;		e:= GetFirstEmailFromOL("feng@wtv12ef.com")
; 		~~ change the current window
;		SwitchMessageWinFromOL()
; 
;-------------------------------------------------------------------------------
SwitchMessageWinFromOL()
{
	global @error_msg
	IfWinExist, ahk_class rctrl_renwnd32 
	{
		WinActivate  
		WinMaximize  
	}else{
		@error_msg := "[" . A_ThisFunc . "] Err - Message window was not found! `r`n"
		Gosub, ErrorLabel
	}
}
