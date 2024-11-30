import json
from datetime import datetime as dt
from tabulate import tabulate
from pathlib import Path
import os

class Customer:

    def append_cust_table(self, cust_table):
        cust_table[self.cust_id] = {'title':self.title, 'first_name': self.first_name,
                                        'last_name': self.last_name ,
                                        'company_name ': self.company_name,
                                        'email' : self.email,
                                        'phone_number':self.phone_number ,
                                        'street':self.street,
                                        'number':self.number,
                                        'postal_code':self.postal_code ,
                                        'city':self.city ,
                                        'country':self.country,
                                        'customer type':self.customer_type,
                                        'account_creation_date':self.account_creation_date,
                                        'VAT':self.VAT ,
                                        'Bank account':self.Bank_account}
        return(cust_table)

    def __init__ (self):
        self.cust_id = 'id'
        self.title = 'title'
        self.first_name = 'first_name'
        self.last_name = 'last_name'
        self.company_name = 'company_name'
        self.email = 'email'
        self.phone_number = 'phone_number'
        self.street = 'street'
        self.number = 'number'
        self.postal_code = 'postal_code'
        self.city = 'city'
        self.country = 'country'
        self.customer_type = 'customer type:'
        self.account_creation_date = dt.today().date().strftime('%d-%m-%y')
        self.VAT = 'VAT:'
        self.Bank_account = 'Bank account:'


    def define_cust_id(self,cust_table):
        id_list = list(cust_table.keys())
        id_list.sort(reverse = True, key = int)
        self.cust_id = str(int(id_list[0]) + 1)
    
    def encode_cust(self, title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account):
        self.title = title
        self.first_name = first_name
        self.last_name = last_name
        self.company_name = company_name
        self.email = email
        self.phone_number = phone_number
        self.street = street
        self.number = number
        self.postal_code = postal_code
        self.city = city
        self.country = country
        self.customer_type = customer_type
        self.account_creation_date = account_creation_date
        self.VAT = VAT
        self.Bank_account = Bank_account

    def import_cust(self, id, cust_table):
        self.cust_id = id
        try:
            self.title = cust_table[id]['title']
        except:
            self.title = 'Mr'
        self.first_name = cust_table[id]['first_name']
        self.last_name = cust_table[id]['last_name']
        self.company_name = cust_table[id]['company_name']
        self.email = cust_table[id]['email']
        self.phone_number = cust_table[id]['phone_number']
        self.street = cust_table[id]['street']
        self.number = cust_table[id]['number']
        self.postal_code = cust_table[id]['postal_code']
        self.city = cust_table[id]['city']
        self.country = cust_table[id]['country']
        try:
            self.customer_type = cust_table[id]['customer_type:']
        except:
            self.customer_type = 'prof'
        self.account_creation_date = cust_table[id]['account_creation_date']
        self.VAT = cust_table[id]['VAT']
        self.Bank_account = cust_table[id]['Bank account']

    def display_cust(self):
        b = [[i,j] for i,j in self.__dict__.items()]
        return(tabulate(b, headers=['Key', 'Value']))

    def __str__ (self):
        return ('{0}, {1}, {2}'.format(self.first_name, self.last_name, self.company_name))
    
    def export_cust(self):
        a = {self.cust_id : {
            'first_name': self.first_name, 
            'last_name': self.last_name, 
            'email': self.email, 
            'phone_number': self.phone_number , 
            'company_name': self.company_name,
            'customer_type': self.customer_type, 
            'account_creation_date': self.account_creation_date, 
            'VAT': self.VAT, 
            'Bank account': self.Bank_account, 
            'street': self.street, 
            'city': self.city, 
            'postal_code': self.postal_code, 
            'country': self.country, 
            'number': self.number, 
            'titre': self.title}}
        return(a)


from textual.app import App, ComposeResult
from textual.widgets import Input


from textual.app import App
from textual.widgets import Button, Header, Input, Label

from textual.app import App
from textual.widgets import Footer

class MyApp(App):
    BINDINGS = [("b", "bell", "Ring")]

    def compose(self):
        yield Footer()

    def action_bell(self):
        self.bell()
        self.mount(Label("Ring!"))

if __name__ == "__main__":
    app = MyApp()
    app.run()