'''

Created on Feb 9, 2016




@author: Imran Ahmed

'''




from nested_dict import nested_dict
import random
#import win32com.client
import xlrd
import sqlite3
import time




class VirtualCoffee:

    def __init__(self):

        self.filepath = "./virtual_coffee_test_data.xlsx"
        
        #conn = sqlite3.connect('vcp.db')
        #conn.execute('''CREATE TABLE PAIRS (
        #ID INT PRIMARY KEY     NOT NULL, 
        #email_emp_1    VARCHAR(50),
        #email_emp_2    VARCHAR(50),
        #date            DATE     NOT NULL);''')        
        #conn.close()
        #return None

    

    def read_file(self):

        """

        Open and read an Excel file

        """

        book = xlrd.open_workbook(self.filepath)

        # get the first worksheet

    

        first_sheet = book.sheet_by_index(0)

        max_cols = first_sheet.ncols

        max_rows = first_sheet.nrows

        min_ct = 0

        #print(max_rows)

        emp_dict = nested_dict();

        dict_head = dict();

        #read a row slice

        for emp_no in range(min_ct, max_rows):  

            sliced_row = first_sheet.row_slice(rowx=emp_no,

                                               start_colx=min_ct,

                                               end_colx=max_cols)

            

            head_ct = 0

            if emp_no == 0:

                for head in sliced_row:

                    dict_head[head_ct] = head.value

                    head_ct += 1

            else:

                for defcell in sliced_row:

                    emp_dict[emp_no][dict_head[head_ct]] = defcell.value

                    head_ct += 1

        return emp_dict




    def pair(self, emp_dictionary):

        max_pairs = int(len(emp_dictionary.keys()) / 2) + 1

        #print(max_pairs)

        emp_dictionary.pop(0, None)

#'''

        rand_num = 0

        pairs_srno = []

        for pairing in range(1, max_pairs):

        #rand_num_2=randint(1,len(emp_dictionary.keys()))i

            #print("In pairing loop")

            couple = []

            for cnt in (0, 1):

                rand_num = random.choice(list(emp_dictionary.keys()))

                emp_dictionary.pop(rand_num, None)

                couple.insert(cnt, rand_num)
                
                #print(couple)

                pairs_srno.insert(pairing-1, couple)
                                
        return pairs_srno
    
    def emailbody(self, name1, wiki1, name2, wiki2):

        emailbody = '<p>Hi! <br/> You have been invited for Virtual Coffe Program. <br/><a href="%s">%s</a> and <a href="%s">%s</a></p>' % (wiki1, name1, wiki2, name2)         

        return emailbody

    def email(self, toEmail, emailSubject, emailBody):

        olMailItem = 0x0

        #obj = win32com.client.Dispatch("Outlook.Application")

        #newMail = obj.CreateItem(olMailItem)

        #newMail.Subject = emailSubject

        #newMail.HTMLBody  = emailBody
        
        #newMail.SentOnBehalfOfName = "virtual.coffee@lrn.com" 
        
        #newMail.To = toEmail        

        #newMail.Attachments.Add(attachment)                

        #newMail.Send()    
        
    def insert(self,name1,name2,date):        
        conn = sqlite3.connect('vcp.db', timeout=100)        
        conn.execute('INSERT INTO PAIRS (email_emp_1,email_emp_2,date) VALUES (?,?,?)',(name1,name2,date))
        conn.commit()
        conn.close()            



if __name__ == "__main__":

    #path = "C:/Users/Imran.Khan/Desktop/Python_text/virtual_coffee_test_data.xlsx"

    #print("Hi")

    vcp = VirtualCoffee()

    emp_hash = vcp.read_file()

    pair_ids = vcp.pair(emp_hash)

    restore_emp_hash = vcp.read_file()
    
    #print(pair_ids)
    date = time.strftime("%d/%m/%Y")
    
    for couple in pair_ids:
       
        #print(restore_emp_hash[couple[0]]['Email'],restore_emp_hash[couple[0]]['Profile'],"<PAIR>",restore_emp_hash[couple[1]]['Email'],restore_emp_hash[couple[1]]['Profile'])        
        name1 = restore_emp_hash[couple[0]]['Name']  #Name 1 
        wiki1 = restore_emp_hash[couple[0]]['Profile'] #Wiki 1 
        email1 = restore_emp_hash[couple[0]]['Email'] #Email 1 
        
        name2 = restore_emp_hash[couple[1]]['Name'] #Name 2         
        wiki2 = restore_emp_hash[couple[1]]['Profile'] #Wiki 2
        email2 = restore_emp_hash[couple[1]]['Email'] #Email 2   
        
        vcp.insert(email1,email2,date)
        
        emailbody = vcp.emailbody(name1, wiki1, name2, wiki2)
        
        toEmail = email1 + ';' + email2
        
        vcp.email(toEmail, 'Invitation for Virtual Coffe', emailbody)
    #print(pair_ids)
    
    #print(len(emp_dictionary.keys()))    
