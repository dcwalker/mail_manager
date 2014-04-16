require 'net/imap'
require 'net/smtp'
require 'openssl'
require 'yaml'
require 'htmlentities'
require "uri"
require "base64"

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
      send_mail_drop_messages(imap_authenticated_connection, smtp_config, m.name, days_until_reminder, "#{email_prefix} #{short_name}:", reminder_email_prefix)
    end
  end

  imap_authenticated_connection.select(mailbox_name)

  unless days_until_reminder.nil?
    unless days_until_reminder.is_a? Array
      days_until_reminder = [days_until_reminder]
    end
    days_until_reminder.each do |days|
      reminder_date = Date.today - days.to_i
      puts " Sending reminder tasks for messages before #{reminder_date.strftime("%d-%b-%Y")}"
      imap_authenticated_connection.uid_search(["KEYWORD", "SentToMailDrop", "NOT", "KEYWORD", "SentReminderToMailDrop", "NOT", "DELETED", "BEFORE", reminder_date.strftime("%d-%b-%Y")]).each do |message_id|
        imap_authenticated_connection.uid_store(message_id, "+FLAGS", ["SentReminderToMailDrop"])
        imap_authenticated_connection.uid_store(message_id, "-FLAGS", ["SentToMailDrop"])
      end
    end
  end

  puts " Messages in \'#{mailbox_name}\' folder to send to MailDrop:"
  imap_authenticated_connection.uid_search(["NOT", "KEYWORD", "SentToMailDrop", "NOT", "DELETED"]).each do |message_id|
    flags = imap_authenticated_connection.uid_fetch(message_id, "FLAGS")[0].attr["FLAGS"]
    envelope = imap_authenticated_connection.uid_fetch(message_id, "ENVELOPE")[0].attr["ENVELOPE"]
    print "  (#{message_id}) "
    puts "#{envelope.from[0].name}: \t#{envelope.subject} -> #{flags.join(", ")}"

    headers_to_forward = "CONTENT-TYPE CONTENT-TRANSFER-ENCODING"
    headers_to_forward = imap_authenticated_connection.uid_fetch(message_id,"BODY[HEADER.FIELDS (#{headers_to_forward})]")[0].attr["BODY[HEADER.FIELDS (#{headers_to_forward})]"]
    headers_to_forward = headers_to_forward.strip unless headers_to_forward.nil?
    content_type = imap_authenticated_connection.uid_fetch(message_id,"BODY[HEADER.FIELDS (CONTENT-TYPE)]")[0].attr["BODY[HEADER.FIELDS (CONTENT-TYPE)]"]
    boundary = content_type.match(/.*boundary=(.*)$/) unless content_type.nil?
    boundary = boundary[1].gsub!(/\"/, "") if boundary
    email_body = imap_authenticated_connection.uid_fetch(message_id,"BODY[TEXT]")[0].attr["BODY[TEXT]"]

    subject = envelope.subject
    subject.slice!(/Subject:\s/i)
    subject.gsub!(/=\?.*?\?Q\?/, "")
    subject.gsub!(/\?=/, "")
    subject.gsub!(/\?/, "=3f")
    if flags.include?("SentReminderToMailDrop")
      subject = "#{reminder_email_prefix} #{email_prefix} #{subject}"
    else
      subject = "#{email_prefix} #{subject}"
    end

    message_id_url = envelope.message_id.sub(/<(.*)\>/i, "message:%3C\\1%3E <message:%3C\\1%3E>")
    plain_text = "#{message_id_url}\r\n\r\n"
    html_text = "#{message_id_url}<br>\r\n<br>\r\n"
    if boundary.nil?
      if content_type.include?("html")
        email_body = "#{html_text}#{email_body}"
      else
        email_body = "#{plain_text}#{email_body}"
      end
    else
      if !email_body.match(/--.*?Content-Type: text\/html;.*?Content-Transfer-Encoding: base64.*?\r\n\r\n/im).nil?
        html_text = Base64.encode64(html_text)
      end
      email_body.sub!(/(--#{boundary}.*?Content-Type: text\/plain.*?\r\n\r\n)/im, "\\1#{plain_text}")
      email_body.sub!(/(--#{boundary}.*?Content-Type: text\/html;.*?\r\n\r\n)/im, "\\1#{html_text}")
    end

    message_string = <<END_OF_MESSAGE
#{headers_to_forward}
Date: #{envelope.date}
From: #{smtp_config["from_address"]}
To: #{smtp_config["to_address"]}
Subject: =?UTF-8?Q?#{URI.escape(HTMLEntities.new.decode(subject)).gsub(/%/, "=")}?=

#{email_body}
END_OF_MESSAGE

    smtp = Net::SMTP.new(smtp_config["server"], 587)
    smtp.enable_starttls
    smtp.start(smtp_config["helo_domain"], smtp_config["username"], get_password_from_keychain(smtp_config["server"], smtp_config["username"]), :plain) do |smtp|
      smtp.send_message message_string, smtp_config["from_address"], smtp_config["to_address"]
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
