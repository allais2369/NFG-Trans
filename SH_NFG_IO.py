# -*- coding: utf-8 -*-
"""
Created on Thu May 17 16:44:05 2018

@author: Nathaniel Bass

Written for Susannah's House for import/export of data between
Quickbooks and Network for Good.
This script is provided free of charge and may not be used in commercial
software, in full or in part, without the written permission of its author.
All other uses are cool, as long as:
this disclaimer is included;
the author is credited;
all the bits and pieces of the derived work are still free to use in the
above-described manner.

© 2018 Nathaniel Bass

"""

import re
import csv
import datetime

quickbooks_exported_names = [
    'Aging',
    'Open Balance',
    'Qty',
    'U/M',
    'Sales Price',
    'Debit',
    'Credit',
    'Amount',
    'Trans #',
    'Type',
    'Entered/Last Modified',
    'Last modified by',
    'Date',
    'Rep',
    'Sales Tax Code',
    'Clr',
    'Estimate Active',
    'Billing Status',
    'Split',
    'Print',
    'Paid',
    'Pay Meth',
    'Name Account #',
    'Memo',
    'Pay Period Begin Date',
    'Pay Period End Date',
    'Check #',
    'Grant Program',
    'Newsletter',
    'Item',
    'Item Description',
    'SSN/Tax ID',
    'Payroll Item',
    'Income Subject To Tax',
    'Wage Base',
    'Wage Base (Tips)',
    'Account',
    'Class',
    'Adj',
    'P. O. #',
    'Name',
    'Source Name',
    'Name E-Mail',
    'Name Address',
    'Name Street1',
    'Name Street2',
    'Name City',
    'Name State',
    'Name Zip',
    'Name Contact',
    'Name Phone #',
    'Name Fax #',
    'Ship Date',
    'Deliv Date',
    'FOB',
    'Via',
    'Terms',
    'Due Date',
    'Billed Date',
    'Paid Through',
    'Driver\'s License',
    'Prayer Team',
    'Volunteer',
    'Informal Greeting',
    'Household Name (The _Family)',
    'Formal Greeting',
    'Num',
    'Tax Table Version',
    'User Edit?',
    'Calculated Amount',
    'Amount Difference',
    'S. O. #',
    'Account Type',
    'Tax Line',
    'WC Rate',
    'Exp. Mod.',
    'WC Code',
    'State',
    'Action',
    'Title Number',
    'Backordered',
    'Avg Days to Pay',
    'Paid Date',
    'Ship To City',
    'Ship To Address 1',
    'Ship To Address 2',
    'Ship To State',
    'Ship Zip',
    'Other 1',
    'Other 2',
    'Preferred Delivery Method',
    'Paycheck Date',
    'Current Rate',
    'Previous Rate',
    '% Change',
    'Pay Period',
    'Notes',
    'Amount Paid'
    ]

network_for_good_to_import_names = [
    'amount', # Must be a positive number with no punctuation other than a decimal point. (i.e. 12.34 (good), $200,292 (bad)). Must be greater than 0 unless the payment type is 'in_kind', in which case, it must be 0
    'donated_at', # The date of the donation, the date format should be 'yyyy-mm-dd' (i.e. 2015-02-09) or 'yyyy-mm-dd hh:mm:ss' (i.e. 2015-02-09 11:20:00)
    'payment_method', # Must be ach, cash, check, credit_card, direct, gift_card, givecard, in_kind, invoice, match, other, paypal, payroll or square.  Anything else will be converted to "other".
    'description', # Any text note to be attached to the donation
    'payment_desc', # Donation check number and date (i.e #2034, 02-06-2016)
    'designation', # The name of a program area or similar that the funds are to be directed to. If the name is not found in the database, a new designation will be created
    'campaign', # The name of a fundraising campaign that the donation will be attached to. If no campaign is found with the same name, a new campaign will be created.
    'tribute_name', # Name of person to be honored/memorialized
    'tribute_type', # Must be honor, memorial, on_behalf or not_specified
    'tribute_notification_method', # Must be none, mail, email, or not_specified
    'acknowledged', # Use any of the following to indicate that the donation has been acknowledged: 'true', 't', 1. Use any of the following to indicate that the donation is not acknowledged: 'false', 'f', 0, or blank
    'fair_market_value', #
    'pledge_id', # Must be the id of an existing pledge for the contact
    'internal_id', # The numeric internal ID of the individual record. Only empty fields will be updated. New data will not overwrite existing data. To update donation records, please use the update_donation importer.
    'external_user_id', # Use this value to ensure that multiple donation records get attributed to a single donor. All records with the same external_user_id will be attached to one donor record using the information in the first row containing that external_user_id
    'user_notes', # Notes are displayed on the contact's profile
    'first_name', #
    'middle_name', #
    'last_name', #
    'full_name', #
    'salutation', # Mr, Mrs, Ms, Dr, etc
    'suffix', #
    'email', # Must be a valid email address
    'work_email', #
    'home_phone', #
    'mobile_phone', #
    'work_phone', #
    'full_address', #
    'address', #
    'address_2', #
    'city', #
    'state', #
    'zip_code', #
    'country', # If other address fields are included, and country is left blank, it will be set to 'US'
    'formal_greeting', #
    'household_name', #
    'informal_greeting', #
    'full_work_address', #
    'work_address', #
    'work_address_2', #
    'work_city', #
    'work_state', #
    'work_zip_code', #
    'work_country', #
    'employer', #
    'job_title', #
    'date_of_birth', # YYYY-MM-DD format (ex: 1984-07-14)
    'gender', # Must be either m or f (m = male, f = female)
    'photo_url', # Include the full URL to the photo. NFG will store the original photo as well as a cropped 300x300 thumbnail which gets displayed in the app. Add photos to existing contacts using internal_id but existing images will not get changed or removed.
    'head_of_household', #
    'dob_year', # 4 digit year (ex: 1984)
    'dob_month', # 1-12 (ex: 7)
    'dob_day', # 1-31 (ex: 14)
    'receive_emails', # True or false.  Defaults to true
    'custom_field_fax', #  Fax # (phone_number)
    'custom_field_prayer-team', # Prayer Team (select) N, Y
    'custom_donation_field_item', # Item (textbox)
    'group_50_mile_radius', #
    'group_3/24/18_info_event_rsvps', #
    'group_board_of_directors', #
    'group_volunteers', #
    'group_newsletter', #
    ]

quickbooks_date_format = "%m/%d/%Y"
network_for_good_donated_at_format = '%Y-%m-%d %H:%M:%S'
network_for_good_birthdate_format = '%Y-%m-%d'
network_for_good_date_format = '%Y-%m-%d'
email_regex = re.compile(r"[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$")

network_for_good_payment_methods = [
    "ach",
    "cash",
    "check",
    "credit_card",
    "direct",
    "gift_card",
    "givecard",
    "in_kind",
    "invoice",
    "match",
    "other",
    "paypal",
    "payroll",
    "square",
    ]

tribute_types = ["honor", "memorial", "on_behalf", "not_specified"]

tribute_notification_methods = ["none", "mail", "email", "not_specified"]

quickbooks_exported_filename = "QBExport 05-14-18.CSV"
quickbooks_exported_filename = "QBExport05-24-18.CSV"
network_for_good_to_import_filename = "NFG Import {}.csv".format(
    datetime.datetime.today().strftime('%Y%m%d_%H%M')
    )

def network_for_good_phone_number_format(phone_number):
    if len(phone_number) == 7:
        print "fix me! {}".format(phone_number)
        phone_number = "865-{}-{}".format(phone_number[:3],phone_number[3:])
        print "okay, but we had to guess the area code! {}\n".format(phone_number)
        return phone_number
    if len(phone_number) == 10:
        print "fix me! {}".format(phone_number)
        phone_number = "{}-{}-{}".format(phone_number[:3],phone_number[3:6],phone_number[6:])
        print "okay! {}\n".format(phone_number)
        return phone_number
    if len(phone_number) == 12:
        if phone_number[-5] == "-":
            print "good phone number {}".format(phone_number)
            return phone_number
    if phone_number == "":
        return phone_number
    print "\nAw, nutz... {} is no good\n".format(phone_number)
    return phone_number

def print_names_in_quickbooks_export():
    with open(network_for_good_to_import_filename, "r") as f:
        g = 0
        for row in f.readlines():
            print len(row.split(","))
            g += 1
            if g == 15:
                break
        f.seek(0)
        names = f.readline()
        names = names.split('\n')[0]
        names = [names.split(","),]
        names.append(f.readline().split("\n")[0].split(","))
        names.append(f.readline())
    print names

def parse_quickbooks_exported_data():
    quickbooks_row_temp = []
    quickbooks_exported_data = []
    with open(quickbooks_exported_filename, 'rb') as f:
        for row in csv.reader(f):
            if any(e for e in row[1:-2] if e != ''):
                for entry in row:
                    quickbooks_row_temp.append(entry)
                quickbooks_exported_data.append(quickbooks_row_temp)
                continue
            # quickbooks_row_temp = [row[0],]
            quickbooks_row_temp = []
    quickbooks_exported_names_dic = {name : i \
        for i,name in enumerate(quickbooks_exported_data[0]) \
        }
    return quickbooks_exported_names_dic, quickbooks_exported_data[1:]

def parse_network_for_good_to_import_data():
    qb_n, qb_d = parse_quickbooks_exported_data()
    network_for_good_to_import_data = [network_for_good_to_import_names,]
    for entry in qb_d:
        row_temp = []
        for heading in network_for_good_to_import_names:
            if heading == 'amount':
                amount = float(entry[qb_n["Amount"]])
                row_temp.append("{:.2f}".format(amount))
            elif heading == 'donated_at':
                date = datetime.datetime.strptime(
                    entry[qb_n["Date"]],
                    quickbooks_date_format
                    )
                row_temp.append(
                    datetime.datetime.strftime(
                        date,
                        network_for_good_date_format
                        )
                    )
            elif heading == 'payment_method':
                pay_meth = entry[qb_n["Pay Meth"]].lower()
                if pay_meth not in network_for_good_payment_methods:
                    print "Payment method \"{}\" \
                        converted to \"other\"".format(pay_meth)
                row_temp.append(pay_meth)
            elif heading == 'description':
                row_temp.append(entry[qb_n["Memo"]])
            elif heading == 'payment_desc':
                row_temp.append(entry[qb_n["Check #"]])
            elif heading == 'designation':
                row_temp.append('')
            elif heading == 'campaign':
                row_temp.append(entry[qb_n["Class"]])
            elif heading == 'tribute_name':
                row_temp.append('')
            elif heading == 'tribute_type':
                row_temp.append('')
            elif heading == 'tribute_notification_method':
                row_temp.append('')
            elif heading == 'acknowledged':
                row_temp.append('')
            elif heading == 'fair_market_value':
                row_temp.append('')
            elif heading == 'pledge_id':
                row_temp.append('')
            elif heading == 'internal_id':
                row_temp.append(entry[qb_n["Num"]])
            elif heading == 'external_user_id':
                row_temp.append('')
            elif heading == 'user_notes':
                row_temp.append('')
            elif heading == 'first_name':
                row_temp.append('')
            elif heading == 'middle_name':
                row_temp.append('')
            elif heading == 'last_name':
                row_temp.append('')
            elif heading == 'full_name':
                row_temp.append(entry[qb_n["Name"]])
            elif heading == 'salutation':
                row_temp.append('')
            elif heading == 'suffix':
                row_temp.append('')
            elif heading == 'email':
                email = entry[qb_n["Name E-Mail"]]
                if email == "":
                    row_temp.append('')
                elif email_regex.match(email):
                    row_temp.append(email)
                else:
                    print "Email {} not valid, omitted.".format(email)
            elif heading == 'work_email':
                row_temp.append('')
            elif heading == 'home_phone':
                phone_number = entry[qb_n["Name Phone #"]]
                phone_number = network_for_good_phone_number_format(phone_number)
                row_temp.append(phone_number)
            elif heading == 'mobile_phone':
                row_temp.append('')
            elif heading == 'work_phone':
                row_temp.append('')
            elif heading == 'full_address':
                row_temp.append('')
            elif heading == 'address':
                row_temp.append(entry[qb_n["Name Street1"]])
            elif heading == 'address_2':
                row_temp.append(entry[qb_n["Name Street2"]])
            elif heading == 'city':
                row_temp.append(entry[qb_n["Name City"]])
            elif heading == 'state':
                row_temp.append(entry[qb_n["Name State"]])
            elif heading == 'zip_code':
                row_temp.append(entry[qb_n["Name Zip"]])
            elif heading == 'country':
                row_temp.append('')
            elif heading == 'formal_greeting':
                row_temp.append(entry[qb_n["Formal Greeting"]])
            elif heading == 'household_name':
                row_temp.append(entry[qb_n["Household Name (The _Family)"]])
            elif heading == 'informal_greeting':
                row_temp.append(entry[qb_n["Informal Greeting"]])
            elif heading == 'full_work_address':
                row_temp.append('')
            elif heading == 'work_address':
                row_temp.append(entry[qb_n["Ship To Address 1"]])
            elif heading == 'work_address_2':
                row_temp.append(entry[qb_n["Ship To Address 2"]])
            elif heading == 'work_city':
                row_temp.append(entry[qb_n["Ship To City"]])
            elif heading == 'work_state':
                row_temp.append(entry[qb_n["Ship To State"]])
            elif heading == 'work_zip_code':
                row_temp.append(entry[qb_n["Ship Zip"]])
            elif heading == 'work_country':
                row_temp.append('')
            elif heading == 'employer':
                row_temp.append(entry[qb_n["Source Name"]])
            elif heading == 'job_title':
                row_temp.append('')
            elif heading == 'date_of_birth':
                row_temp.append('')
            elif heading == 'gender':
                row_temp.append('')
            elif heading == 'photo_url':
                row_temp.append('')
            elif heading == 'head_of_household':
                row_temp.append('')
            elif heading == 'dob_year':
                row_temp.append('')
            elif heading == 'dob_month':
                row_temp.append('')
            elif heading == 'dob_day':
                row_temp.append('')
            elif heading == 'receive_emails':
                row_temp.append('')
            elif heading == 'custom_field_fax':
                row_temp.append(entry[qb_n["Name Fax #"]])
            elif heading == 'custom_field_prayer-team':
                row_temp.append('')
            elif heading == 'custom_donation_field_item':
                row_temp.append(entry[qb_n["Memo"]])
            elif heading == 'group_50_mile_radius':
                row_temp.append('')
            elif heading == 'group_3/24/18_info_event_rsvps':
                row_temp.append('')
            elif heading == 'group_board_of_directors':
                row_temp.append('')
            elif heading == 'group_volunteers':
                row_temp.append('')
            elif heading == 'group_newsletter':
                row_temp.append('')
            else:
                row_temp.append('')
                print heading
        network_for_good_to_import_data.append(row_temp)
    delete_rows = []
    row_0 = []
    for i, row in enumerate(network_for_good_to_import_data):
        if row == row_0:
            delete_rows.append(i)
        row_0 = row
    for i in delete_rows[::-1]:
        network_for_good_to_import_data.pop(i)
    print "Hi "
    print "{} duplicate entries were deleted. ".format(len(delete_rows))
    return network_for_good_to_import_data

def network_for_good_to_import_write():
    network_for_good_to_import_filename = "NFG Import {}.csv".format(
        datetime.datetime.today().strftime('%Y%m%d_%H%M')
        )
    network_for_good_to_import_data = parse_network_for_good_to_import_data()
    with open(network_for_good_to_import_filename, 'wb') as f:
        writer = csv.writer(f)
        writer.writerows(network_for_good_to_import_data)
    print "{} entries were written to \"{}\". ".format(
        len(network_for_good_to_import_data),
        network_for_good_to_import_filename
        )
