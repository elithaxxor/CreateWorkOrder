import os, subprocess
import datetime as dt
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import docx

class InvoiceAutomation:
    def __init__(self):
       # self.create_invoice = None
        self.root = tk.Tk() # Create a new Tkinter window
        self.root.title('Invoice Automation')
        self.root.geometry('500x800')


        # Create the button to create the invoice
        self.create_button = tk.Button(self.root, text='Create Invoice', command=self.create_invoice)

        # Create the labels and entries for the invoice automation form
        self.partner_label = tk.Label(self.root, text='Partner:')
        self.partner_entry = tk.Entry(self.root)

        self.partner_street_label = tk.Label(self.root, text='Partner Street:')
        self.partner_street_entry = tk.Entry(self.root)

        self.partner_zip_city_country_label = tk.Label(self.root, text='Partner Zip, City, Country:')
        self.partner_zip_city_country_entry = tk.Entry(self.root)

        self.invoice_number_label = tk.Label(self.root, text='Invoice Number:')
        self.invoice_number_entry = tk.Entry(self.root)

        self.invoice_date_label = tk.Label(self.root, text='Invoice Date:')
        self.invoice_date_entry = tk.Entry(self.root)

        self.service_description_label = tk.Label(self.root, text='Service Description:')
        self.service_description_entry = tk.Entry(self.root)

        self.service_amount_label = tk.Label(self.root, text='Service Amount:')
        self.service_amount_entry = tk.Entry(self.root)


        self.service_single_amount_label = tk.Label(self.root, text='Service Single Amount:')
        self.service_single_amount_entry = tk.Entry(self.root)

        self.payment_terms_label = tk.Label(self.root, text='Payment Terms:')
        self.payment_terms_entry = tk.Entry(self.root)

        self.payment_method_label = tk.Label(self.root, text='Payment Method:')


        # Create the dropdown for the payment method
        self.payment_method_entry = tk.Entry(self.root)
        self.payment_methods = {'Credit Card', 'Check', 'Wire Transfer'}
        self.payment_methods = tk.StringVar(self.root)
        self.payment_methods.set('Wire Transfer')
        self.payment_method_dropdown = tk.OptionMenu(self.root, self.payment_methods, "Credit Card", "Check", "Wire Transfer")


        # Pack the widgets  in the window

        padding_options = {'fill': tk.X, 'expand': True, 'padx': 5, 'pady': 5}

        self.partner_label.pack(padding_options)
        self.partner_entry.pack(padding_options)

        self.partner_street_label.pack(padding_options)
        self.partner_street_entry.pack(padding_options)

        self.partner_zip_city_country_label.pack(padding_options)
        self.partner_zip_city_country_entry.pack(padding_options)

        self.invoice_number_label.pack(padding_options)
        self.invoice_number_entry.pack(padding_options)

        self.invoice_date_label.pack(padding_options)
        self.invoice_date_entry.pack(padding_options)

        self.service_description_label.pack(padding_options)
        self.service_description_entry.pack(padding_options)

        self.service_amount_label.pack(padding_options)
        self.service_amount_entry.pack(padding_options)

        self.service_single_amount_label.pack(padding_options)
        self.service_single_amount_entry.pack(padding_options)

        self.payment_terms_label.pack(padding_options)
        self.payment_terms_entry.pack(padding_options)

        self.payment_method_label.pack(padding_options)

    def create_invoice(self):
        pass

def create_invoice(): # Create the invoice
    pass

