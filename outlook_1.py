__author__ = "onkar.kubal"
__date__ = "$8 Feb, 2016 4:32:52 PM$"

if __name__ == "__main__":
    print ("Hello World")
    
import win32com.client

class InviteCoffe:
    
    def sendEmail(self,toEmail,toSubject,toBody):        
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = toSubject
        newMail.HTMLBody  = toBody
        newMail.To = toEmail        
        #newMail.Attachments.Add(attachment)                
        newMail.Send()                
        
x = InviteCoffe() # Object of the class

x.sendEmail('onkar.kubal@lrn.com','Virtual Coffee Program Follow-up Demo','<p>Hello World</p><p> SO AM I!!!</p><p> <b>OnCar</b>') # Send the Email