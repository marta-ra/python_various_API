# import sendgrid
# import os
import requests
import json

class ApiError(Exception):
    """An API Error Exception"""

    def __init__(self, status):
        self.status = status

    def __str__(self):
        return "APIError: status={}".format(self.status)

data_send = {
  "list_ids": [
    "cf123211-8093-440c-9feb-2251792f41dd"
  ],
  "contacts": [
    {
      "alternate_emails": [
        "test@test.net"
      ],
      "country": "Canada",
      "email": "test@test.net",
      "first_name": "firstname",
      "last_name": "lastname",
      "state_province_region": "Q1",
       "custom_fields": {
        "e86_T": "tel_us"
        }
      }
  ]
}


headersAPI = {
    'Content-Type': 'application/json',
    'Authorization': ''
}
response = requests.get('https://api.sendgrid.com/v3/marketing/contacts', headers=headersAPI, verify=True)
api_response = response.json()
print(response.json)
send = requests.put('https://api.sendgrid.com/v3/marketing/contacts/put', headers=headersAPI, data=data_send, verify=True)
print(send.json)
print(send.headers)
# if response.status_code != 200:
#     # This means something went wrong.
#     raise ApiError('GET /tasks/ {}'.format(response.status_code))
# for todo_item in response.json():
#     print('{} {}'.format(todo_item['id'], todo_item['summary']))