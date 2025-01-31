import logging
from mt_api.connection import get_connection
import time

logger = logging.getLogger().getChild("MAIL")

def send_mail(subject, body, recipient):
    """Sends an email using the database inbuilt function. Done using SQL query"""
    t_sql_command = """
    EXEC msdb.dbo.sp_send_dbmail @profile_name = ?,
                        @recipients = ?,
                        @subject = ?,
                        @body = ?;
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(t_sql_command, ("MIE Notifications", recipient, subject, body))

        logger.info(f"Email command executed. Recipient -> {recipient}. SENT STATUS PENDING")

    time.sleep(1)

    with get_connection() as conn:
        query = "SELECT IDENT_CURRENT('msdb.dbo.sysmail_allitems')"
        cursor.execute(query)

        pk = cursor.fetchone()
        if pk:
            pk = pk[0]
        sent_status_query = "SELECT sent_status FROM msdb.dbo.sysmail_allitems WHERE mailitem_id = ?"
        cursor.execute(sent_status_query, pk)
        status = cursor.fetchone()[0] 
        if status == "sent":
            logger.info("Eamil Sent. SENT STATUS CONFIRMED")
        elif status == "failed":
            logger.error(f"Email failed. recipient: {recipient}")
        else:
            logger.critical("Unknown error. Message not delivered")

