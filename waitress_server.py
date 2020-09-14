from waitress import serve
import upload
import logging
logger = logging.getLogger('waitress')
logger.setLevel(logging.INFO)
serve(upload.app, host='10.27.6.245', port=5000, max_request_body_size=3221225472)