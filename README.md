# xlsSendEmailViaOutlookSemiAutomatically
According to the list on the Excel file, send an e-mail via Outlook semi-automatically.   
Excel上のリストに従って半自動でOutlookを使ってメールを送る。


## Overview
A tool that sends e-mails semi-automatically and continuously while checking visually according to the information in the list on Excel. Allow attachment of different attachments for each destination.  
Excel上のリストの情報に従ってメールを半自動（見て確認しながら）で連続的に送信するツール。宛先ごとに異なった添付ファイルを添付できるようにする。

## Usage
### step_1
Input some information as following in the 'mail' worksheet.

1. Subject.
1. Signature.
1. Body of the e-mail.
1. List of the e-mail address of the recipient.
1. List of the e-mail address for the carbon copy
1. List of the name of the recipient.  
    The sting [Name of the recipient] in the body of the e-mail will be replaced name of the recipient.
1. List of the relative path to the attachment.

Image of the excel sheet.<br>
<img src="https://github.com/okagen/xlsSendEmailViaOutlookSemiAutomatically/blob/master/img01.png?raw=true" width="600">

### step_2
Click the [send mail] button, then some e-mail windows will appear.   

e-mail windows shown after click the [send mail] button.<br>
<img src="https://github.com/okagen/xlsSendEmailViaOutlookSemiAutomatically/blob/master/img02.png?raw=true" width="600">

Change the following line if you don't need to check the e-mail visually.  
  - beore changing.
~~~
objMailItem.Display
~~~
  - afer changing.
~~~
objMailItem.Send
~~~

## Requirement
1. Excel 2013
