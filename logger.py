#!/usr/bin/env python3

'''Configure logging for uncaught exceptions'''

import logging
import sys


class ExtraNewlineFormatter(logging.Formatter):
    '''Wrapper for logging.Formatter to add newlines after log entries'''
    def format(self, record):
        original = super().format(record)
        return f"{original}\n"

def configure_logger():
    '''Set up the root logger for logging uncaught exceptions'''
    logger = logging.getLogger()

    if not logger.hasHandlers():
        logger.setLevel(logging.ERROR)
        file_handler = logging.FileHandler("exceptions.log", mode='a')
        formatter = ExtraNewlineFormatter('%(asctime)s')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    def handle_exception(exc_type, exc_value, exc_traceback):
        '''Function to wrap uncaught exceptions so they can be logged'''
        logger.error("", exc_info=(exc_type, exc_value, exc_traceback))
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    # overwrite internal function for uncaught exceptions
    sys.excepthook = handle_exception
