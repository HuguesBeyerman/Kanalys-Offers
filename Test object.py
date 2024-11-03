

import json
from datetime import datetime as dt
import click
from tabulate import tabulate
from pathlib import Path
import os


class Offer:

    def __init__ (self):  
        self.offer_id = ''b.count(dt.today().date().strftime('%d-%m-%y')) + 1''
        self.Customer = ''
        self.Building': ''
        self.Bom: {}, 
        self.SOW: {},
        self.Path: 'offer_path', 
        self.Status: 'in approval',
        self.Created: dt.today().date().strftime('%d-%m-%y'),
        self.Schema: ''

    def define_offer_id(self, offer_table):
        a = list(offers_table.keys())
        b = [i[:8] for i in a]
        self.offer_id = b.count(dt.today().date().strftime('%d-%m-%y')) + 1)
    

class Customer:

    def append_cust_table(self, cust_table):
        cust_table[self.customer_id] = {'title':self.title, 'first_name': self.first_name,
                                        'last_name': self.last_name ,
                                        'company_name ': self.company_name,
                                        'email' : self.email,
                                        'phone_number':self.phone_number ,
                                        'street':elf.street,
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

        #return(self)

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
        print(cust_table[id])
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

        def __str__ (self):
            return ('{0}, {1}, {2}'.format(self.first_name, self.last_name, self.company_name)


class Building:
    
    def __init__ (self, Street, Number, CP, City, Country, cust_id ):       
        self.build_id = ''
        self.street = Street
        self.number = Number
        self.postal_code= CP
        self.city= City
        self.country= Country
        self.manager = cust_id
        self.account_creation_date = dt.today().date().strftime('%d-%m-%y')

    def define_build_id(self,build_table):
        id_list = list(build_table.keys())
        id_list.sort(reverse = True, key = int)
        self.build_id = str(int(id_list[0]) + 1)

    def encode_cust_address(self, Street, Number, CP, City, Country, cust ):
        self.street = Street
        self.number = Number
        self.postal_code= CP
        self.city= City
        self.country= Country
        self.manager = cust.cust_id
 
    def import_build(self, id, build_table):
        self.build_id = id
        self.street = build_table[id]['street']
        self.number = build_table[id]['number']
        self.postal_code = build_table[id]['postal_code']
        self.city = build_table[id]['city']
        self.country = build_table[id]['country']
        self.manager = build_table[id]['manager']
        self.account_creation_date = build_table[id]['account_creation_date']

    def __str__ (self):
        return ('{0}, {1}, {2}, {3}'.format(self.street, self.number, self.postal_code, self.city)

    def append_building_table(self, build_table):
        build_table[self.build_id] = {
                            'street'= self.street
                            'number' = self.number 
                            'number' = self.postal_code
                            'city' = self.city
                            'country' = self.country
                            'manager' = self.manager
                            'account_creation_date' = self.account_creation_date}  
        return(build_table)

class Product:
    def __init__ (self, description, unit_selling_price, unit_cost):       
        id_list = list(prod_table.keys())
        id_list.sort(reverse = True, key = int)
        self.prod_id = str(int(id_list[0]) + 1)
        self.prod_id = ''
        self.description = description
        self.unit_selling_price = unit_selling_price
        self.unit_cost = unit_cost

    def define_prod_id(self,prod_table):
        id_list = list(prod_table.keys())
        id_list.sort(reverse = True, key = int)
        self.prod_id = str(int(id_list[0]) + 1)

    def encode_prod(self,description, unit_selling_price, unit_cost):
        self.description = description
        self.unit_selling_price = unit_selling_price
        self.unit_cost = unit_cost

    def import_prod(self, id, prod_table):
        self.prod_id = id
        self.description = prod_table[id]['description']
        self.unit_selling_price = prod_table[id]['unit_selling_price']
        self.unit_cost = prod_table[id]['unit_cost']

    def __str__ (self):
        return ('{0}, {1}, {2}'.format(self.description, self.unit_selling_price, self.unit_cost)

    def append_prod_table(self, prod_table):
            prod_table[self.prod_id] = {
                    'description'= self.description
                    'unit_selling_price' = self.unit_selling_price
                    'unit_cost' = self.unit_cost}
            return(prod_table)

class Bom:
    def create_bom_entity(self, product, description, unit_selling_price, qty, discount, line_price):
        seld.bom_id = ''
        self.prod_id = product
        self.description = description
        self.unit_selling_price = unit_selling_price
        self.qty = qty
        self.discount = discount
        self.line_price = line_price

    def define_bom_id(self,bom_table):
        self.bom_id = len(bom_table) + 1

    def encode_bom(self,prod, qty, discount):
        self.prod_id = prod.prod_id
        self.description = prod.description
        self.unit_selling_price = prod.unit_selling_price
        self.qty = qty
        self.discount = discount
        self.line_price = self.qty * self.unit_selling_price* (1 - self.discount/100)
        
    def append_bom_table(self, bom_table):          
        bom_table[self.bom_id] = {
                    'prod_id'= self.prod_id
                    'description' = self.description
                    'unit_selling_price' = self.unit_selling_price
                    'discount' = self.discount
                    'qty' = self.qty
                    'line_price' = self.line_price}
        return(bom_table)


def open_table():
    with open("json/customers.json", 'r') as file:
        cust_table = json.load(file)
    #products and service    
    with open("json/sewer_products_services.json", 'r') as file:
        prod_table = json.load(file)
    # SOW paragraps
    with open("json/paragraphs.json", 'r') as file:
        parag_table = json.load(file)
    #buildings
    with open("json/buildings.json", 'r') as file:
        build_table = json.load(file)
    #offers    
    with open('json/offers.json', 'r') as file:
        offers_table = json.load(file)
        
    return(cust_table, build_table, prod_table, offers_table, parag_table)

def save_json(cust_table, build_table, offers_table, parag_table, prod_table):
    # Saving the list as a JSON file
    
    with open('json/customers.json', 'w') as file:
        json.dump(cust_table, file, indent=4)
    
    with open('json/offers.json', 'w') as file:
        json.dump(offers_table, file, indent=4)
        
    with open('json/buildings.json', 'w') as file:
        json.dump(build_table, file, indent=4)

    with open('json/sewer_products_services.json', 'w') as file:
        json.dump(prod_table, file, indent=4)

    with open('json/paragraphs.json', 'w') as file:
        json.dump(parag_table, file, indent=4)

    print('Fichiers sauvés')
    
# Clear Screen
def clear_screen():
    input('\nPress Enter to continue.')
    if os.name == 'nt':  # For Windows
        os.system('cls')
    else:  # For macOS and Linux
        os.system('clear')

@click.command()
#Start
def start():
    cust_table, build_table, prod_table, offers_table, parag_table = open_table()
    
    x = True
    while x:
        click.clear()   
        click.echo(tabulate([['Kanalys Offers']], tablefmt="heavy_grid"))
        choice = click.prompt('1.Encode new offer, 2.Modify existing offer, 3.Approve offer, 4.Quit ', type=click.Choice([1, 2, 3, 4])
        
        if choice == 1:
            cust_table, build_table, prod_table, offers_table, parag_table = create_offer(cust_table, build_table, prod_table, offers_table, parag_table)

        ''''''elif choice == 2:
            #publish the offers selection   
            query = input('Search on customer name :')
            if query == None:
                query = 'All'
            try:    
                date_beg= dt.strptime(input("Earliest date (format 'dd-mm-yy') "),'%d-%m-%y')
            except ValueError:
                date_beg = dt.strptime('01-01-00', '%d-%m-%y')

            try:
                date_end = dt.strptime(input("Latest date (format 'dd-mm-yy') "), '%d-%m-%y')
            except ValueError:
                date_end = dt.today()
                
            offer_list = create_list_offers(offers_table, cust_table, s=query, date_beg=date_beg, date_end=date_end)
            
            #select offer to modify
            offer_id = input('\nWhich offer ID do you want to modify? ')
            y = True

            #check if offer_id exist and create new offer if not in approval  
            if offer_id in offer_list:
                
                if offers_table[offer_id]['Status'] != 'in approval':
                    new_offer_id = create_offer_nbr(offers_table)
                    print('This offer has already been sent. An new one will be created with ID nbr ' + new_offer_id)
                    offers_table['new_offer_id'] = offers_table['new_offer_id']
                else:
                    new_offer_id = offer_id
                    
                cust_table, build_table, prod_table, offers_table, parag_table = modify_offer(new_offer_id, cust_table, build_table, prod_table, offers_table, parag_table)

            else:
                print('This ID is not in the Offer list')
                y = False

        elif choice == '3':
            approve_offer(cust_table, build_table, prod_table, offers_table, parag_table)'''
            
        else:
            save_json(cust_table, build_table, offers_table, parag_table, prod_table)
            os.abort()
            x = False

'offer creation process'
def create_offer(cust_table, build_table, prod_table, offers_table, parag_table):
    offer = Offer()
    offer.define_offer_id(offers_table)
    cust_table, cust = CLI.select_client(cust_table, offer)
    build_table, build = CLI.select_building(build_table, cust, offer)
    bom_table = {}
    prod_table, bom_table = select_bom(prod_table, bom_table, offer)
    content = {}
    parag_table, content = select_content(parag_table, content, offer_id)
    path_schema = crop_image(offer_id)
    z=True
    while z:
        
        bom_table = create_bom_table(bom)
        image = Image.open(path_schema)
        image.show()
        display_offre(bom_table, content, offer_id, cust_table, cust_id, parag_table, build_table, build_id)
        z, cust_id, cust_table, build_id, build_table, bom, content, prod_table, parag_table, path_schema = change_offer(cust_id, cust_table, build_id, build_table, bom, content, prod_table, parag_table, offer_id, path_schema)

    bom_table = create_bom_table(bom)
    offer_path, path_schema = create_letter(cust_table, cust_id, bom, content, build_table, build_id, offer_id, bom_table, parag_table, path_schema)
    offers_table = create_offer_table(offers_table, offer_id, cust_id, build_id, bom, content, offer_path, path_schema)
    return(cust_table, build_table, prod_table, offers_table, parag_table)

def start_test():
    open_table()
    cust = Customer()
    cust.define_cust_id(cust_table)
    title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account = CLI.encode_cust(cust)
    cust.modify(title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account)
    cust.display_cust()



class CLI:
    def encode_cust(cust):
        click.echo('Encode your customer\n-------------------\n')
        title = click.prompt('title (default: {}) '.format(cust.title), default=cust.title)
        first_name = click.prompt('first_name (default: {}) '.format(cust.first_name), default=cust.first_name)
        last_name = click.prompt('last_name (default: {}) '.format(cust.last_name), default=cust.last_name)
        company_name = click.prompt('company_name (default: {}) '.format(cust.company_name), default=cust.company_name)
        email = click.prompt('email (default: {}) '.format(cust.email), default=cust.email)
        phone_number = click.prompt('tel (default: {}) '.format(cust.phone_number), default=cust.phone_number)
        street = click.prompt('street (default: {}) '.format(cust.street), default=cust.street)
        number = click.prompt('number (default: {}) '.format(cust.number), default=cust.number)
        postal_code = click.prompt('postal_code (default: {}) '.format(cust.postal_code), default=cust.postal_code)
        city = click.prompt('city (default: {}) '.format(cust.city), default=cust.city)
        country = click.prompt('country (default: {}) '.format(cust.country), default=cust.country)
        customer_type = click.prompt('customer type: (default: {}) '.format(cust.customer_type), default=cust.customer_type, type=click.Choice(['Prof', 'Privé']))
        account_creation_date = cust.account_creation_date
        VAT = click.prompt('VAT: (default: {}) '.format(cust.VAT), default=cust.VAT)
        Bank_account = click.prompt('Bank account: (default: {}) '.format(cust.Bank_account), default=cust.Bank_account)
        return(title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account)

    def encode_build(build):
        click.echo('Encode your building\n-------------------\n')
        street = click.prompt('street (default: {}) '.format(build.street), default=build.street)
        number = click.prompt('number (default: {}) '.format(build.number), default=build.number)
        postal_code = click.prompt('CP (default: {}) '.format(build.postal_code), default=build.postal_code)
        city= click.prompt('city (default: {}) '.format(build.city), default=build.city)
        country = click.prompt('country (default: {}) '.format(build.country), default=build.country)
        return(street, number, postal_code, city, country)
        
    
    def select_client(cust_table, offer):
        cust = Customer()
        click.clear()    
        click.echo(tabulate([['Client - offer ID ' + offer.offer_id]], tablefmt="simple_grid"))
        choice = click.prompt('\n1. Existing Customer 2.Encode new customer 3.Search Customer 4.Abort ', click.choice=[1,2,3,4])
        if choice == 1:
            cust_id = click.prompt('\nSélectionner le Customer ID: ')           
            if cust_id in list(cust_table.keys()):
                cust.import_cust(cust_id, cust_table)
                click.echo('Result: ' + cust)
            else:
                click.echo(cust_id + ' is not in customer list')
                continue  
            click.pause('Enter to continue')
            i= False
            
        elif choice == 2:
            cust.define_cust_id
            title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account = CLI.encode_cust(cust)
            cust.modify(title, first_name, last_name, company_name, email, phone_number, street, number, postal_code, city, country, customer_type, account_creation_date, VAT, Bank_account)
            cust_table = cust.append_cust_table(cust_table)
            click.echo(cust.display_cust())
            click.pause('Enter to continue')
            i = False
            
        elif choice == 3:
            CLI.find_cust(cust_table)
            
        else:
            if click.confirm('Do you confirm that you want to quit the prg? All previously encoded offers will be saved.'):
                save_json(cust_table, build_table, offers_table, parag_table, prod_table)
                os.abort()
            else:
                continue

        return(cust_table, cust)
        
    ''' query customer list: allows searching by name or company name using find_cust.'''
    def find_cust(cust_table):
        click.echo('\nRecherche client existant\n--------------------------')
        s = click.prompt('\nCustomer query: ')
        l = []
        for i,j  in cust_table.items():
            if (s in (j.get('last_name') or '')) or (s in (j.get('company_name') or '')) or (s in (j.get('first_name') or '')):
                l.append([i, j['first_name'], j['last_name'], j['company_name']])
        if len(l) > 0:
            click.echo('\n', tabulate(l, headers =['ID', 'Nom', 'Prenom', 'Cny']))
        else:
            click.echo('\n no candidate')

    def select_building(build_table, cust, offer):   
        build = Building()
        click.clear()
        click.echo(tabulate([['Building - Offer ID :' + offer.offer_id]], tablefmt="simple_grid"))
    
        no_build =  CLI.query_building(build_table, cust)
            
        if not no_build:
            exist = click.prompt('\n1. Existing Building 2.Encode new building 3.Abort ', click.choice=[1,2,3])

            x = True
            while x:
                if exist == 1:                    
                        build_id = click.echo('\nBuilding ID: ')
                        
                        if build_id in list(build_table.keys()):
                            build.import_build(build_id, build_table)
                            click.echo('Result : ' +  build)
                            x = CLI.match_build_cust(x, cust, build_table, build)
                        else:
                            click.echo(build_id + ' is not in building list')
                            continue
                        
                elif exist == '2':
                        street, number, postal_code, city, country = encode_build(build)
                        build.encode_cust_address(street, number, postal_code, city, country, cust )
                        build_table = build.append_building_table(build_table)
                        x = False
        
                else:
                    if click.confirm('Do you confirm that you want to quit the prg? All previously encoded offers will be saved.'):
                        save_json(cust_table, build_table, offers_table, parag_table, prod_table)
                        os.abort()
                    else:
                        continue
                        
            else :
                    street, number, postal_code, city, country = encode_build(build)
                    build.encode_cust_address(street, number, postal_code, city, country, cust )
                    x = False
    
    
        return(buidl_table, build)


    
    '''check if building is managed by client'''
    def match_build_cust(x, cust.cust_id , build_table, build):
        #print(type(build_table[build_id]['manager']), type(cust_id))
        if build_table[build.build_id]['manager'] != cust.cust_id:
            click.echo('The building is not managed by this client!')
        else:
            x = False
        return(x)

    ''' Takes a customer ID and searches the building data for buildings managed by that customer.
    Displays a list of buildings or indicates no buildings found.'''
    def query_building(build_table, cust):
        l = []
        no_build = False
        for i,j  in build_table.items():
            if cust.cust_id == j['manager']:
                l.append([i, j['street'], j['number'], j['postal_code'], j['city']])
        if len(l) > 0:
            click.echo('\nListe des immeubles gérés par cust nbr '+ cust. cust_id + '\n--------------------------------------------')
            click.echo((tabulate(l, tablefmt="plain")))
            
        else:
            click.echo("\nPas d'immeuble connu")
            no_build = True
        return(no_build)

    
    '''Provides options to add existing items, create new products, erase items, exit BOM creation or abort'''
    def select_bom(prod_table, bom_table, offer):
        
        build = Building()
        click.clear()
        click.echo(tabulate([['Bill of Materials- Offer ID :' + offer.offer_id]], tablefmt="simple_grid"))
    
        x = True
        while x:
            choice = click.prompt('\n1. Existing Building 2.Encode new building 3.Abort ', click.choice=[1,2,3])

            if choice == 1:
                prod_table, bom_table, bom_id = select_prod_item(prod_table, bom_table)
                
            elif choice == '2':
                prod_table, bom, bom_id = creation_prod_item(prod_table, bom, bom_id)
                
            elif choice == '3':
                #display bom items table
                l=[]
                for i, j in bom.items():
                    l.append([i, j['description'], j['unit_selling_price'], j['qty'], j['discount']])
                if len(l) > 0:
                    print('Items in the offer Bill of Materials')
                    print('\n', tabulate(l, headers =['Bom ID', 'Description', 'Price', 'Qty', 'Disc' ]))
                else:
                    print('\n no content')

                #query item to erase
                a = input('\nItem to erase: ')
                b = [str(c) for c in list(bom.keys())]
                if a in b:
                    del(bom[a])
                    print('Item is erased.')
                else:
                    print('Item not in the list.')
                    
            elif choice == '4':
                find_prod(prod_table)
                    
            elif choice == '5' and len(bom) == 0:
                print('\nThere must be at least on item in the Bill of Material')
                
            elif choice == '6':
                choix  = input('Do you confirm that you want to quit the prg? All previously encoded offers will be saved. (Y,N): ')
                if choix == 'Y':
                    save_json(cust_table, build_table, offers_table, parag_table, prod_table)
                    os.abort()
                elif choix == 'N':
                    continue
                else:
                    print('Encode Y or N')
            
            else:
                x = False
    
                    
        return(prod_table, bom)   


        
    '''Allows adding existing products from a product list to the BOM.
    Takes user input for product ID and calculates the line price considering quantity and discount.'''
    def select_prod_item(prod_table, bom_table):
    
        a = click.prompt('\nProd ID: ')
        prod = Product()
        bom = Bom()
        
        if a in list(prod_table.keys()):
            bom.define_bom_id()
            prod.import_prod(a, prod_table)
            click.echo('Result : ' +  prod)
            bom.qty = click.prompt('Qty: ', click.INT)
            bom.discount = click.prompt('Discount (%): ', click.IntRange(min_open = 0, max_open = 100)))               
            bom.line price = bom.qty * bom.unit_selling_price * (1 - bom.discount/100)          
            bom.encode_bom(prod, qty, discount)
            prod_table = bom.append_prod_table(prod_table)
                
        else:
            print(a + ' is not in product list')
        
    return(prod_table, bom_table)


if __name__ == "__main__":
   start()

    def append_building_table(self, build_table):
        build_table[self.build_id] = {
                            'street'= self.street
                            'number' = self.number 
                            'number' = self.postal_code
                            'city' = self.city
                            'country' = self.country= cust[cust_id]['country']
                            'manager' = self.manager = cust[cust_id]
                            'account_creation_date' = self.account_creation_date}  
        return(build_table)
