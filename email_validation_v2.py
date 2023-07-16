import http.client
import mimetypes
import json
import ssl
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename


print("Select the CSV file containing the list of emails")

Tk().withdraw() 
csv_email_file = askopenfilename() 

email_list=pd.read_csv(csv_email_file, sep=",", header=None)

###
# for testing
#list = ["umanggpta@gmail.com","ugtalk2@gmail.com","asflafsnjk6728@gmail.com","kranthi2193@gmail.com"]
#email_list=pd.DataFrame([list])
###

email_result = pd.DataFrame()

context = ssl._create_unverified_context()
conn = http.client.HTTPSConnection("api.eva.pingutil.com", context=context)
payload = ''
headers = {}


#start_value = 10001
#end_value = 25000
#step_value = 1000
#j = 0

print("\nThe total number of emails in the CSV file is:",len(email_list.columns))


start_value = int(input("\nEnter start_value (enter '1' to start from the first email in the CSV):"))

#end_value = int(input("Enter end_value (enter the total number of emails to verify):"))
end_value = len(email_list.columns)

print("\n'step_value' is the number of emails after which the result will be saved in excel file before proceeding to next set of emails")
step_value = int(input("Enter step_value (type '0' to not split the output):"))
if step_value == 0:
    step_value = end_value


for i in range(start_value-1,end_value):
        
    email=email_list.at[0,i]

    conn.request("GET", "/email?email={}".format(email), payload, headers)
    res = conn.getresponse()
    data = res.read()

    data = json.loads(data.decode("utf-8"))
    is_valid=data[['data'][0]]['deliverable']
    email_result = pd.concat([email_result, pd.DataFrame([[email,is_valid]])], ignore_index=True)

    print("veryfying email:",i,email,is_valid)
    if (i+1)%step_value == 0:
            email_result.to_excel("email_valid_{}-{}.xlsx".format(i-step_value+2,i+1))
            #j=i
            email_result = pd.DataFrame()

