library(RDCOMClient)
library(readxl)

outlook_app <- COMCreate("Outlook.Application")
# Create a search object to search the mail by subject title
outlookNameSpace <- outlook_app$GetNameSpace("MAPI")
mailbox <- outlookNameSpace$Folders()

# Key in the subject title here
search <- outlook_app$AdvanceSearch(
    "Inbox",
    "urn:schema:httpmail:subject = 'Data for today'" # In this example, the subject title for the email is Data for today
)

Sys.sleep(5)

results <- search$Results()
email <- results$Item(1)
attachment_file <- tempfile()
email$Attachment(1)$SaveAsFile(attachment_file)

# Read data
data <- read_excel(attachment_file)