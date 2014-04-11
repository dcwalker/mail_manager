require 'net/imap'
require 'net/smtp'
require 'openssl'
require 'mail'
require 'yaml'
require 'htmlentities'

Net::IMAP::add_authenticator('PLAIN', Net::IMAP::PlainAuthenticator)
Net::IMAP::add_authenticator('LOGIN', Net::IMAP::LoginAuthenticator)

def get_password_from_keychain(service_name, account_name)
  output = `security find-generic-password -s #{service_name} -a #{account_name} -w 2>&1`
  if output.empty?
    puts "\n"
    puts " Empty string returned from keychain."
    puts " You may need to unlock your keychain first: security unlock-keychain"
    exit 0
  elsif output.include? "The specified item could not be found in the keychain."
    puts "\n"
    puts " The specified item could not be found in the keychain."
    puts " To add a keychain entry for this login use: security add-generic-password -s #{service_name} -a #{account_name} -w <password>"
    exit 0
  else
    return output.strip
  end
end

def verify_smtp_connection(smtp_config)
  smtp = Net::SMTP.new(smtp_config["server"], 587)
  smtp.enable_starttls
  smtp.start(smtp_config["helo_domain"], smtp_config["username"], get_password_from_keychain(smtp_config["server"], smtp_config["username"]), :plain) do |smtp|
  end
end


def mark_as_seen (imap_authenticated_connection, mailboxes)
  print " Marking messages as \"Seen\" in: "
  mailboxes.each do |mailbox|
    next unless imap_authenticated_connection.list('', mailbox)
    print "#{mailbox} "
    imap_authenticated_connection.select(mailbox)
    count = 0
    imap_authenticated_connection.uid_search(["NOT", "SEEN"]).each do |message_id|
      imap_authenticated_connection.uid_store(message_id, "+FLAGS", [:Seen])
      count =+ 1
    end
    print "(#{count}) "
  end
  puts "\n"
end

def expunge_mailboxes (imap_authenticated_connection, mailboxes)
  print " Expunging messages in: "
  mailboxes.each do |mailbox|
    next unless imap_authenticated_connection.list('', mailbox)
    print "#{mailbox} "
    imap_authenticated_connection.select(mailbox)
    imap_authenticated_connection.expunge
  end
  puts "\n"
end

def delete_messages (imap_authenticated_connection, rules)
  rules.each do |rule|
    next unless imap_authenticated_connection.list('', rule["in_mailbox"])
    puts " Marking messages 'Deleted' in #{rule["in_mailbox"]} #{rule["field"]} #{rule["address"]}"
    imap_authenticated_connection.select(rule["in_mailbox"])
    imap_authenticated_connection.search([rule["field"].upcase, rule["address"]]).each do |message_id|
      imap_authenticated_connection.store(message_id, "+FLAGS", [:Deleted])
    end
  end
end

def send_mail_drop_messages (imap_authenticated_connection, smtp_config, mailbox_name, days_until_reminder=nil, email_prefix="", reminder_email_prefix="")
  if not imap_authenticated_connection.list('', mailbox_name)
    imap_authenticated_connection.create(mailbox_name)
  elsif imap_authenticated_connection.list("", "#{mailbox_name}/*")
    imap_authenticated_connection.list(mailbox_name, "*").each do |m|
      short_name = m.name.split(m.delim).last
      email_prefix = "#{email_prefix} #{short_name}:"
      send_mail_drop_messages(imap_authenticated_connection, smtp_config, m.name, days_until_reminder, email_prefix, reminder_email_prefix)
    end
  end

  imap_authenticated_connection.select(mailbox_name)

  unless days_until_reminder.nil?
    reminder_date = Date.today - days_until_reminder.to_i
    puts " Sending reminder tasks for messages before #{reminder_date.strftime("%d-%b-%Y")}"
    imap_authenticated_connection.uid_search(["KEYWORD", "SentToMailDrop", "NOT", "KEYWORD", "SentReminderToMailDrop", "NOT", "DELETED", "BEFORE", reminder_date.strftime("%d-%b-%Y")]).each do |message_id|
      imap_authenticated_connection.uid_store(message_id, "+FLAGS", ["SentReminderToMailDrop"])
      imap_authenticated_connection.uid_store(message_id, "-FLAGS", ["SentToMailDrop"])
    end
  end

  puts " Messages in \'#{mailbox_name}\' folder to send to MailDrop:"
  imap_authenticated_connection.uid_search(["NOT", "KEYWORD", "SentToMailDrop", "NOT", "DELETED"]).each do |message_id|
    flags = imap_authenticated_connection.uid_fetch(message_id, "FLAGS")[0].attr["FLAGS"]
    envelope = imap_authenticated_connection.uid_fetch(message_id, "ENVELOPE")[0].attr["ENVELOPE"]
    print "  (#{message_id}) "
    puts "#{envelope.from[0].name}: \t#{envelope.subject} -> #{flags.join(", ")}"
    mail = Mail.read_from_string imap_authenticated_connection.uid_fetch(message_id,'RFC822')[0].attr['RFC822']

    email_prefix = "#{reminder_email_prefix} #{email_prefix}" if flags.include?("SentReminderToMailDrop")
    maildrop_mail = Mail.new do
      to      smtp_config["to_address"]
      from    smtp_config["from_address"]
      subject "Subject: #{HTMLEntities.new.decode(email_prefix)} #{HTMLEntities.new.decode(mail.subject)}"
    end

    text_part_body = ""
    html_part_body = ""
    content_type = ""
    if mail.multipart?
      mail.parts.each do |part|
        if part.content_type.start_with? "text/plain"
          text_part_body += part.body.to_s
          content_type = part.content_type
        elsif part.content_type.start_with? "text/html"
          html_part_body += part.body.to_s
          content_type ||= part.content_type
        elsif !part.filename.nil?
         maildrop_mail.attachments[part.filename] = { :content => part.body.to_s }
        else
          text_part_body += "could not parse: #{part.inspect}"
        end
      end
    else
      text_part_body = mail.body
      content_type = mail.content_type
    end

    text_part = Mail::Part.new do
      content_type = (content_type.nil? || content_type.empty?) ? "text/plain; charset=UTF-8" : content_type
      content_type content_type
      text = text_part_body.empty? ? html_part_body : text_part_body
      body "message:<#{mail.message_id}>\n\n#{text}\n"
    end
    maildrop_mail.text_part = text_part
    maildrop_mail.charset = mail.charset unless mail.charset.nil?

    smtp = Net::SMTP.new(smtp_config["server"], 587)
    smtp.enable_starttls
    smtp.start(smtp_config["helo_domain"], smtp_config["username"], get_password_from_keychain(smtp_config["server"], smtp_config["username"]), :plain) do |smtp|
      smtp.send_message maildrop_mail.to_s, smtp_config["from_address"], smtp_config["to_address"]
    end

    imap_authenticated_connection.uid_store(message_id, "+FLAGS", ["SentToMailDrop"])
    imap_authenticated_connection.uid_store(message_id, "+FLAGS", [:Flagged])
  end
end

def cleanup_old_maildrop_messages(imap_authenticated_connection)
  print " Cleaning flags on previous maildrop messages in: "
  imap_authenticated_connection.list('', '*').each do |mailbox|
    next if mailbox.name.include? "OmniFocus tasks"
    print "#{mailbox.name} "
    imap_authenticated_connection.select(mailbox.name)
    count = 0
    imap_authenticated_connection.uid_search(["KEYWORD", "SentToMailDrop"]).each do |message_id|
      imap_authenticated_connection.uid_store(message_id, "-FLAGS", ["SentToMailDrop"])
      imap_authenticated_connection.uid_store(message_id, "-FLAGS", ["SentReminderToMailDrop"])
      imap_authenticated_connection.uid_store(message_id, "-FLAGS", [:Flagged])
      count =+ 1
    end
    print "(#{count}) "
  end
  puts "\n"
end

config = YAML.load_file(File.join(File.expand_path(File.dirname(__FILE__)), "config.yaml"))

verify_smtp_connection(config["smtp"])

config["accounts"].each do |account|
  puts "Processing #{account["description"]}"
  imap = Net::IMAP.new(account["imap_server"], :ssl => true)
  imap.login(account["login"], get_password_from_keychain(account["imap_server"], account["login"]))

  mark_as_seen(imap, account["mark_as_seen"]) if account.has_key?("mark_as_seen")

  delete_messages(imap, account["delete_messages"]) if account.has_key?("delete_messages")
  send_mail_drop_messages(imap, config["smtp"], "OmniFocus tasks", account["days_until_reminder"], account["email_prefix"], account["reminder_email_prefix"])
  cleanup_old_maildrop_messages(imap)
  expunge_mailboxes(imap, account["expunge_mailboxes"]) if account.has_key?("expunge_mailboxes")

  imap.logout
end