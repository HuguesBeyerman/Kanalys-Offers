
#import uuid  # For generating unique IDs

class Customer:
    customer_table = {}  # Dictionary to store all customer instances
    
    def __init__(self, title, first_name, last_name, company_name, email, tel, street, number, postal_code, city, country):
        #self.customer_id = str(uuid.uuid4())  # Unique ID per instance
        self.title = title
        self.first_name = first_name
        self.last_name = last_name
        self.company_name = company_name
        self.email = email
        self.tel = tel
        self.street = street
        self.number = number
        self.postal_code = postal_code
        self.city = city
        self.country = country

        # Add the instance to the customer table using customer_id as the key
        Customer.customer_table[self.customer_id] = self

    @classmethod
    def query_customers(cls, last_name=None, company_name=None):
        """Query customers by last name or company name."""
        results = []
        for customer in cls.customer_table.values():
            if (last_name and customer.last_name == last_name) or (company_name and customer.company_name == company_name):
                results.append(customer)
        return results





from textual.app import App, ComposeResult
from textual.widgets import Input, Button, ListView, ListItem#, TextArea

class CustomerApp(App):

    def compose(self) -> ComposeResult:
        # Interface for creating a customer
        #yield TextArea("Create a Customer")
        yield Input(placeholder="Title", id="title")
        yield Input(placeholder="First Name", id="first_name")
        yield Input(placeholder="Last Name", id="last_name")
        yield Input(placeholder="Company Name", id="company_name")
        yield Input(placeholder="Email", id="email")
        yield Input(placeholder="Telephone", id="tel")
        yield Input(placeholder="Street", id="street")
        yield Input(placeholder="Number", id="number")
        yield Input(placeholder="Postal Code", id="postal_code")
        yield Input(placeholder="City", id="city")
        yield Input(placeholder="Country", id="country")
        yield Button(label="Create Customer", id="create_button")

        # Interface for querying and displaying customers
        #yield TextArea("Query Customers")
        yield Input(placeholder="Last Name", id="query_last_name")
        yield Input(placeholder="Company Name", id="query_company_name")
        yield Button(label="Search", id="search_button")
        yield ListView(id="results_list")

    def on_button_pressed(self, event) -> None:
        # Handle customer creation
        if event.button.id == "create_button":
            fields = {field.id: field.value for field in self.query("Input")}
            customer = Customer(**fields)
            self.update_results_list([customer])

        # Handle customer search
        elif event.button.id == "search_button":
            last_name = self.query_one("#query_last_name").value
            company_name = self.query_one("#query_company_name").value
            results = Customer.query_customers(last_name=last_name, company_name=company_name)
            self.update_results_list(results)

    def update_results_list(self, customers):
        """Display customer search results in the ListView."""
        results_list = self.query_one("#results_list")
        results_list.clear()
        for customer in customers:
            item = ListItem(f"{customer.first_name} {customer.last_name} - {customer.company_name}")
            item.data = customer
            results_list.append(item)

    def on_list_view_selected(self, event) -> None:
        """Edit customer details when selected from the list."""
        selected_customer = event.item.data
        # For each attribute, populate the inputs to allow editing
        for field_name in selected_customer.__dict__.keys():
            if self.query_one(f"#{field_name}", None):
                self.query_one(f"#{field_name}").value = getattr(selected_customer, field_name)

if __name__ == "__main__":
    app = CustomerApp()
    app.run()
