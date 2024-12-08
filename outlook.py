class InviteCoffe:
    
    def sendEmail(self):
        import win32com.client
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "SUBJECT!!"
        newMail.HTMLBody  = "<p>Hello World</p><p> SO AM I!!!</p><p> <b>OnCar</b>"
        newMail.To = "onkar.kubal@lrn.com"
        #attachment1 = "C:/Projects/invite/src/tmp/IE10.JPG"
        #newMail.Attachments.Add(attachment1)
        newMail.Send()                
        
x = InviteCoffe() # Object of the class

x.sendEmail() # Send the Email