import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

#isntall the following package: https://gitea.ksol.io/karolyi/py3-validate-email@v1.0.9
#to install: python -m pip install git+https://gitea.ksol.io/karolyi/py3-validate-email@v1.0.9
from validate_email import validate_email

print("Select the CSV file containing the list of emails")

Tk().withdraw() 
csv_email_file = askopenfilename() 

#csv_email_file = input("Enter the name")

email_list=pd.read_csv(csv_email_file, sep=",", header=None)

email_result = pd.DataFrame()

#start_value = 10000
#end_value = 58449
#step_value = 250
#j = 0

print("\nThe total number of emails in the CSV file is:",len(email_list.columns))
start_value = int(input("\nEnter start_value (enter '1' to start from the first email in the CSV):"))


end_value = len(email_list.columns)

#end_value = int(input("Enter end_value (enter the total number of emails to verify):"))

print("\n'step_value' is the number of emails after which the result will be saved in excel file before proceeding to next set of emails")
step_value = int(input("Enter step_value (type '0' to not split the output):"))
if step_value == 0:
    step_value = end_value 

smtp_email = input("Enter an email that will be used to send a verfication request for the list of emails:")

for i in range(start_value-1,end_value):
        
    is_valid = []
    email=email_list.at[0,i]
    is_valid = validate_email(
        email_address=email,
        check_format=True,
        check_blacklist=True,
        check_dns=True,
        dns_timeout=10,
        check_smtp=True,
        smtp_timeout=10,
        smtp_helo_host='smtp.google.com',
        smtp_from_address=smtp_email,
        smtp_skip_tls=False,
        smtp_tls_context=None,
        smtp_debug=False,
        )
    email_result = pd.concat([email_result, pd.DataFrame([[email,is_valid]])], ignore_index=True)
    #print(email_result)
    print("veryfying email:",i,email,is_valid)
    #print(email)
    #print("Validity:",is_valid)
    if (i+1)%step_value == 0:
            email_result.to_excel("email_valid_{}-{}.xlsx".format(i-step_value+2,i+1))
            #j=i
            email_result = pd.DataFrame()
