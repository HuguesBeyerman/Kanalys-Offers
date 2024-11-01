
import uuid  # For generating unique IDs

class Customer:
    customer_table = {}  # Dictionary to store all customer instances
    
    def __init__(self, title, first_name):#, last_name, company_name, email, tel, street, number, postal_code, city, country):
        self.customer_id = str(uuid.uuid4())  # Unique ID per instance
        self.title = title
        self.first_name = first_name
        #self.last_name = last_name
        #self.company_name = company_name
        #self.email = email
        #self.tel = tel
        #self.street = street
        #self.number = number
        #self.postal_code = postal_code
        #self.city = city
        #self.country = country

        # Add the instance to the customer table using customer_id as the key
        Customer.customer_table[self.customer_id] = {'title':self.title, 'first_name': self.first_name}
        print(Customer.customer_table)

    @classmethod
    def query_customers(cls, last_name=None, company_name=None):
        """Query customers by last name or company name."""
        results = []
        for customer in cls.customer_table.values():
            if (last_name and customer.last_name == last_name) or (company_name and customer.company_name == company_name):
                results.append(customer)
        return results

class CustomerApp:

    def compose(self):
        # Interface for creating a customer
        print("Create a Customer")
        title = input('title ')
        first_name = input('first_name ')
        #last_name = input('last_name ')
        #company_name = input('company_name ')
        #email = input('email ')
        #tel = input('tel ')
        #street = input('street ')
        #postal_code = input('postal_code ')
        #number = input('number ')
        #city = input('city ')
        #country = input('country ')
    
        cust = Customer(title, first_name)#, last_name, company_name, email, tel, street, number, postal_code, city, country)

        print(cust.customer_table)


app = CustomerApp()
app.compose()

