import win32com.client
import datetime
__version__ = 1.2


class Outlook:
    def __init__(self):
        # MailItem documentation:
        # http://msdn.microsoft.com/EN-US/library/microsoft.office.interop.outlook.mailitem_members(v=office.14).aspx
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.mapi = self.outlook.GetNamespace("MAPI")

    def send(self, send,
             to,
             subject,
             body_clear_text,
             body_html='',
             body_format=2,
             from_email='',
             cc='',
             bcc='',
             reply_recipients='',
             flag_text='',
             reminder=False,
             reminder_date_time='',
             importance=1,
             read_receipt=False,
             deferred_delivery_date_time='',
             account_to_send_from='',
             attachments: [] = None):
        # Sends an email. Returns: True/False.
        item = self.outlook.CreateItem(0)    # olMailItem
        item.To = to
        item.Subject = subject
        if body_clear_text:   # default signature is kept if you don't change the .Body property
            item.Body = body_clear_text
        if body_html:
            item.HTMLBody = body_html
        item.BodyFormat = body_format  # 1=plain text, 2=html, 3=rich text
        item.SentOnBehalfOfName = from_email		# sets the "From" field. "" is ok as Outlook just uses default account
        item.CC = cc  # == Recipient1 = sacComObjectReply.Recipients.Add("a@a.com") then Recipient1.Type = 2 #To=1 Cc=2 Bcc=3
        item.BCC = bcc
        item.FlagRequest = flag_text 	# sets follow up flag *for recipients*! very cool. "" is ok
        item.ReplyRecipientNames = reply_recipients		# sac, depreciated, should use item.ReplyRecipients.Add("simon crouch")
        if reminder_date_time:
            item.ReminderTime = reminder_date_time
        item.ReminderSet = reminder
        item.ReadReceiptRequested = read_receipt
        item.Importance = importance  # 2=high, 1=med, 0=low
        if deferred_delivery_date_time:    # '%d/%m/%y %H:%M'
            item.DeferredDeliveryTime = deferred_delivery_date_time
        if account_to_send_from:
            item.SendUsingAccount = item.Application.Session.Accounts.Item(account_to_send_from)
        if attachments:
            [item.Attachments.Add(attachment) for attachment in attachments]
        if send:
            item.Send()
        else:
            item.Display()
        return True

    def outlook_repeat_delay_email(self, to, sub, message, delay_date, repeat_count=1, days_apart=0):
        delay_date = datetime.datetime.strptime(delay_date, '%d/%m/%y %H:%M')
        for _ in range(0, repeat_count):
            date_formatted = delay_date.strftime('%d/%m/%y %H:%M')
            sub = sub + " (" + date_formatted + ")"
            self.send(True, to, sub, message, deferred_delivery_date_time=date_formatted)
            delay_date += datetime.timedelta(days=days_apart)

    def appointments_before_0930(self, days_forward=7):
        """
        Return calendar items which occur before time 0930
        :param days_forward: days ahead to search
        :return yields calendar items
        """
        o_calendar = self.mapi.GetDefaultFolder("9")
        o_items = o_calendar.Items
        o_items.IncludeRecurrences = True
        o_items.Sort("[Start]")

        start_date = datetime.datetime.now().strftime("%d/%m/%Y %H:%M %p")
        end_date = (datetime.datetime.now() + datetime.timedelta(days=days_forward)).strftime("%d/%m/%Y %H:%M %p")
        o_items = o_items.Restrict("[Start] >= '" + start_date + "' AND [Start] <= '" + end_date + "'")

        # Find next week appointments, earlier than 09.30am
        for item in o_items:
            if item.Start.time() < datetime.time(9, 30) \
                    and not item.Start.time() == datetime.time(0, 0)\
                    and not item.Categories == "FreeTime":
                yield item


# Notes:
# To move messages:
# for no in range(messages.Count-1, -1, -1):
#     messages[no].Move(inbox.Folders('Processed'))
