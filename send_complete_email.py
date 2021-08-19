import mail

def main(address: str) -> None:
    mail.send_confirmation_email(address)

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()

    parser.add_argument('-a', '--address', type=str, dest='sender_address', default='', help='email address to send to')
    args = parser.parse_args()

    main(args.sender_address)
