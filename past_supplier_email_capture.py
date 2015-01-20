__author__ = 'Afzal Ahmad'

import time
from django.core.management.base import BaseCommand
import logging

import email as email_module
from trips.models import Email
from trips.wm_parser import WMParser

logger = logging.getLogger(__name__)


class Command(BaseCommand):
  def handle(self, *args, **options):
    """*** One time job to Populate 'supplier_email' in Email table from past superfly emails. ***"""
    start_time = time.time()
    logger.info("Initializing the supplier email populating script ")
    wm_parser = WMParser()
    email_ids = Email.objects.exclude(superfly_raw_email__isnull=True).exclude(
      superfly_raw_email__exact='').values_list('id')
    total_records = email_ids.count()
    skipped = 0
    logger.info("Found %s emails to update." % total_records)
    for id in email_ids:
      email = Email.objects.get(id=id[0])
      logger.info("Updating email_id=%s " % email.id)
      if not email.supplier_email:
        msg = email_module.message_from_string(email.superfly_raw_email)
        email_details = wm_parser.get_email_details_from_msg(msg)
        email.supplier_email = email_details.get("from_email")
        email.save()
        logger.info("Updated email_id=%s " % email.id)
      else:
        skipped += 1
        logger.info("Already updated email_id=%s " % email.id)
    logger.info("Supplier Email populating Job completed successfully")
    logger.info("METRICS: Total running time: %s, Total record: %s, Processed: %s, Skipped: %s" % (
    time.time() - start_time, total_records, total_records - skipped, skipped))
