#!/usr/bin/env python3
"""
Convertit des fichiers vCard (VCF) en CSV compatible Outlook.
Dépendances : vobject (pip install vobject)
"""

import csv
import vobject

VCARD_FILE = 'infile.vcf'
CSV_FILE = VCARD_FILE.rsplit('.', 1)[0] + '.csv'

FIELDNAMES = [
    'First Name',
    'Middle Name',
    'Last Name',
    'Title',
    'Suffix',
    'Nickname',
    'E-mail Address',
    'E-mail 2 Address',
    'E-mail 3 Address',
    'Home Phone',
    'Home Phone 2',
    'Business Phone',
    'Business Phone 2',
    'Mobile Phone',
    'Other Phone',
    'Primary Phone',
    'IMAddress',
    'Job Title',
    'Department',
    'Company',
    'Office Location',
    "Manager's Name",
    "Assistant's Name",
    "Assistant's Phone",
    'Home Street',
    'Home City',
    'Home State',
    'Home Postal Code',
    'Home Country/Region',
    'Personal Web Page',
    'Spouse',
    'Schools',
    'Hobby',
    'Location',
    'Web Page',
    'Birthday',
    'Anniversary',
    'Notes',
    'kind',
    'gender',
    'UID',
]

def get_value_safe(obj, attr, default=''):
    return getattr(obj, attr, default) if obj else default

def convert_one_contact(vcard):
    post = {field: '' for field in FIELDNAMES}

    # KIND (type de contact)
    if hasattr(vcard, 'kind'):
        post['kind'] = vcard.kind.value

    # GENDER
    if hasattr(vcard, 'gender'):
        post['gender'] = vcard.gender.value

    # UID
    if hasattr(vcard, 'uid'):
        post['UID'] = vcard.uid.value

    # NOM
    if hasattr(vcard, 'n'):
        post['Last Name'] = get_value_safe(vcard.n.value, 'family')
        post['First Name'] = get_value_safe(vcard.n.value, 'given')
        post['Middle Name'] = get_value_safe(vcard.n.value, 'additional')
        post['Title'] = get_value_safe(vcard.n.value, 'prefix')
        post['Suffix'] = get_value_safe(vcard.n.value, 'suffix')

    # FN (nom complet, si jamais N absent)
    if not post['First Name'] and hasattr(vcard, 'fn'):
        post['First Name'] = vcard.fn.value

    # ORGANISATION
    if hasattr(vcard, 'org'):
        # org.value est une liste (voir doc vobject)
        post['Company'] = vcard.org.value[0] if vcard.org.value else ''

    # EMAILS
    if hasattr(vcard, 'email_list'):
        for idx, email in enumerate(vcard.email_list[:3]):
            field = 'E-mail Address' if idx == 0 else f'E-mail {idx+1} Address'
            post[field] = email.value

    # TÉLÉPHONES
    if hasattr(vcard, 'tel_list'):
        for tel in vcard.tel_list:
            types = [t.upper() for t in tel.params.get('TYPE', [])]
            number = tel.value
            # Attribution selon le type
            if 'CELL' in types:
                if not post['Mobile Phone']:
                    post['Mobile Phone'] = number
                else:
                    post['Other Phone'] = number
            elif 'HOME' in types:
                if not post['Home Phone']:
                    post['Home Phone'] = number
                else:
                    post['Home Phone 2'] = number
            elif 'WORK' in types:
                if not post['Business Phone']:
                    post['Business Phone'] = number
                else:
                    post['Business Phone 2'] = number
            else:
                if not post['Other Phone']:
                    post['Other Phone'] = number

    # NOTE
    if hasattr(vcard, 'note'):
        post['Notes'] = vcard.note.value

    # ANNIVERSAIRE
    if hasattr(vcard, 'bday'):
        post['Birthday'] = vcard.bday.value

    # AUTRES CHAMPS (ajouter selon besoin)
    # ...

    return post

def main():
    with open(VCARD_FILE, 'r', encoding='utf-8') as input_file:
        raw_input_file = input_file.read()

    contacts = []
    for vcard in vobject.readComponents(raw_input_file):
        contacts.append(convert_one_contact(vcard))

    with open(CSV_FILE, 'w', newline='', encoding='utf-8') as output_file:
        writer = csv.DictWriter(output_file, fieldnames=FIELDNAMES, dialect='excel')
        writer.writeheader()
        writer.writerows(contacts)

if __name__ == '__main__':
    main()