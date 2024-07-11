import os, subprocess
import datetime as dt
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import docx

class InvoiceAutomation:
    def __init__(self):
        self.root = tk.Tk() # Create a new Tkinter window
        self.root.title('Invoice Automation')
        self.root.geometry('500x800')

        # Create the labels and entries for the invoice automation form
        self.partner_label = tk.Label(self.root, text='Partner:')
        self.partner_street_label = tk.Label(self.root, text='Partner Street:')
        self.partner_zip_city_country_label = tk.Label(self.root, text='Partner Zip, City, Country:')
        self.invoice_number_label = tk.Label(self.root, text='Invoice Number:')
        self.invoice_date_label = tk.Label(self.root, text='Invoice Date:')
        self.service_description_label = tk.Label(self.root, text='Service Description:')
        self.service_amount_label = tk.Label(self.root, text='Service Amount:')
        self.service_single_amount_label = tk.Label(self.root, text='Service Single Amount:')
        self.payment_terms_label = tk.Label(self.root, text='Payment Terms:')
        self.payment_method_label = tk.Label(self.root, text='Payment Method:')

        # Create the entries for the invoice automation form
        self.partner_entry = tk.Entry(self.root)
        self.partner_street_entry = tk.Entry(self.root)
        self.partner_zip_city_country_entry = tk.Entry(self.root)
        self.invoice_number_entry = tk.Entry(self.root)
        self.invoice_date_entry = tk.Entry(self.root)
        self.service_description_entry = tk.Entry(self.root)
        self.service_amount_entry = tk.Entry(self.root)
        self.service_single_amount_entry = tk.Entry(self.root)
        self.payment_terms_entry = tk.Entry(self.root)

        # Create the dropdown for the payment method
        self.payment_method_entry = tk.Entry(self.root)
        self.payment_methods = {'Credit Card', 'Check', 'Wire Transfer'}
        self.payment_methods = tk.StringVar(self.root)
        self.payment_methods.set('Wire Transfer')
        self.payment_method_dropdown = tk.OptionMenu(self.root, self.payment_methods, "Credit Card", "Check", "Wire Transfer")

