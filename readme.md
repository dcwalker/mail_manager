I created MailManager as a way to inject the `message:` URL of an email into OmniFocus tasks added via [Mail Drop](http://support.omnigroup.com/omnifocus-mail-drop).

## How does it work?

When you run the script it will check your OS X keychain for login credentials. If an `OmniFocus tasks` folder exists on your mail server, it will send all the messages in that folder to the `to_address` in the configuration (probably your Mail Drop address) and tag the messages in that folder so messages are only sent to OmniFocus once.

## Requirements

Because of the dependance on the OS X keychain for password management, this needs to run on an OS X machine.

Everything uses native Ruby libraries except my dependance on [mikel's mail gem](https://github.com/mikel/mail) (installed with `gem install mail`) as a reliable way to extract the message-id from the mail headers.

## Example configuration

The file should be named `config.yaml` and should go in the same directory as `MailManager`.

    accounts:
      - description: My work email
        login: codemonkey
        imap_server: imap.example.com
        mark_as_seen:
          - Trash
          - Archive
          - Sent Messages
        delete_messages:
          - field: TO
            address: maildrop_user@sync.omnigroup.com
            in_mailbox: Trash
          - field: TO
            address: maildrop_user@sync.omnigroup.com
            in_mailbox: Sent Messages
        expunge_mailboxes:
          - INBOX
          - Sent Messages
      - description: Personal mail
        login: espresso
        imap_server: imap.mail.me.com
        mark_as_seen:
          - Deleted Messages
          - Archive
          - Sent Messages

    smtp:
      to_address: maildrop_user@sync.omnigroup.com
      from_address: alottodo@example.com
      server: smtp.example.com
      helo_domain: example.com
      username: alottodo


### Explanation

`accounts` -- is a list of all the mail accounts to check (work, personal, etc).

`marks_as_seen` _(optional)_ -- sometimes Mail.app gets confused and when I move a message I've read from my inbox to my archive it marks it as unread again, all unread messages in folders in the `mark_as_seen` list as marked as read.

`delete_messages` _(optional)_ -- there are some messages I want deleted right away so they don't clutter my search results when searching in Mail.app.  Messages that match the rules in `delete_messages` are marked deleted when the script runs.

`expunge_mailboxes` _(optional)_ -- Mail.app is lazy about when it actually calls expunge on a folder.  [Status Board](http://panic.com/statusboard/) has a bug where it includes messages marked deleted in its counts of folders.  This option is to work around that.

`smtp` -- The script needs your SMTP credentials so that it can send your tasks to you.

`to_address` -- This is where messages are forwarded to (I use my [Mail Drop](http://www.omnigroup.com/support/omnifocus-mail-drop) address here).

`from_address` -- In today's world of SPAM, this needs to be an address that the SMTP mail server trusts, usually of the same domain.

`server` -- The actual SMTP server, like `smtp.example.com`.

`helo_domain` -- I don't know. For me it's the root domain for my SMTP server.

`username` -- The user to authenticate as when connecting to the SMTP server.