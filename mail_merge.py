import smtplib, ssl, time
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

SMTP_HOST  = 'mail.grf.bg.ac.rs'
SMTP_PORT  = 587
FROM_ADDR  = 'nmilovanovic@grf.bg.ac.rs'
FROM_NAME  = 'Никола Миловановић'
PASSWORD   = open('password.txt').read().strip()


# Важно, промени име наслова задатка
SUBJECT    = 'Путна инфтраструктура, поставка 2. задтака'



BCC_ADDRS  = ['904_22@student.grf.bg.ac.rs']
SHEET      = 'komb_zadatak2'

BATCH_SIZE = 40
SLEEP_SEC  = 30


def make_message(to_addr, body_text):
    msg = MIMEMultipart('alternative')
    msg['From']    = formataddr((Header(FROM_NAME, 'utf-8').encode(), FROM_ADDR))
    msg['To']      = to_addr
    msg['Cc']      = ''
    msg['Subject'] = Header(SUBJECT, 'utf-8')
    msg.attach(MIMEText(body_text.replace(r'\n', '\n'), 'plain', 'utf-8'))
    return msg


def smtp_connect():
    ctx = ssl.create_default_context()
    ctx.set_ciphers('DEFAULT:@SECLEVEL=0')
    s = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
    s.starttls(context=ctx)
    s.login(FROM_ADDR, PASSWORD)
    return s


def main():
    df = pd.read_excel('input.xlsx', sheet_name=SHEET)
    df = df[df['mail'].str.contains('@', na=False)].reset_index(drop=True)
    print(f'Учитано {len(df)} валидних адреса из листа "{SHEET}"')

    batches = [df.iloc[i:i + BATCH_SIZE] for i in range(0, len(df), BATCH_SIZE)]
    print(f'{len(df)} порука → {len(batches)} серија по ≤{BATCH_SIZE}')

    for batch_num, batch in enumerate(batches, 1):
        print(f'\n--- Серија {batch_num}/{len(batches)} ({len(batch)} порука) ---')
        s = smtp_connect()
        for _, row in batch.iterrows():
            to_addr = row['mail'].strip()
            msg = make_message(to_addr, row['msg'])
            rcpt_list = list(set([to_addr] + BCC_ADDRS))
            s.sendmail(FROM_ADDR, rcpt_list, msg.as_string())
            print(f'  ✓ {to_addr}')
        s.quit()
        if batch_num < len(batches):
            print(f'Пауза {SLEEP_SEC}с пре следеће серије...')
            time.sleep(SLEEP_SEC)

    print('\nГотово.')


if __name__ == '__main__':
    main()
