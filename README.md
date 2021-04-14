# Apple-VBA-Email
Modified Ron De Bruin code for sending email upon change in Excel worksheet for Mac. The user can specify a range of a given worksheet that, when changed by a user, will trigger an email to be generated and optionally sent to a certain recipient. Email recipients can be conditonally entered, and the body of the email can contain the contents of the cell value that changed.

Before running you must ensure to embed the RDB apple script into the com.microsoft.excel folder so that Excel can communicate with mail. To do this from the finder window, select go in the top left of the screen and hold down option. The library option will appear and you will select application scripts, and navigate to com.microsoft.excel. Here is where you must save the file.

