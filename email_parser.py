import os
from datetime import datetime
from email import message_from_binary_file
from email.message import Message as EmailMessage
from email.utils import getaddresses, parseaddr, parsedate_to_datetime
from typing import BinaryIO, Dict, List, Optional, Tuple, TypedDict, Union
ENCODINGS = ("ascii", "utf-8", "utf-8-sig", "latin-1", "cp1252")
content_types = {".gif":"image/gif",".doc": "application/msword",".jpg":"image/jpeg",".jpg":"image/jpeg",".png":"image/png",".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".pdf": "application/pdf"}

from mimetypes import guess_type



from extract_msg import Message



class Email(TypedDict):
    subject: str
    body: str
    received_on: Union[datetime, str]
    sender: Dict[str, str]
    author: Dict[str, str]
    to: List[Dict[str, str]]
    cc: List[Dict[str, str]]
    attachment_names: List[str]
    attachment_binaries: List[BinaryIO]
    attachment_types: List[str]




def extract_attachments_from_eml(message: EmailMessage, mail_data: Email) -> None:
    """Extracts attachments from email message.


    Args:
        message (EmailMessage): Message to extract.
        mail_data (Email): Dictionary to modify.


    """
    for part in message.walk():
        if part.get_content_disposition() == "attachment":
            content_type = part.get_content_type()
            bin_data = part.get_payload(decode=True)
            name = part.get_filename()
            if name is not None:
                mail_data["attachment_names"].append(part.get_filename())
                mail_data["attachment_binaries"].append(bin_data)
                mail_data["attachment_types"].append(content_type)




def extract_metadata_from_eml_header(message: EmailMessage, mail_data: Email) -> None:
    """Extracts header metadata from email message.


    Args:
        message (EmailMessage): Message to extract.
        mail_data (Email): Dictionary to modify.


    """
    timestamp = parsedate_to_datetime(message.get("date", ""))
    mail_data["received_on"] = timestamp
    mail_data["subject"] = message.get("subject", "")


    author_name, author_address = parseaddr(message.get("from", ""))
    author_information = {"name": author_name, "smtp_address": author_address}
    mail_data["author"] = author_information
    mail_data["sender"] = author_information


    extract_address_list_from_eml(message, mail_data, "to")
    extract_address_list_from_eml(message, mail_data, "cc")




def extract_address_list_from_eml(
    message: EmailMessage, mail_data: Email, key: str
) -> None:
    """Extracts name and SMTP address from message.


    Args:
        message (EmailMessage): Message to extract.
        mail_data (Email): Dictionary to modify.
        key (str): Key to extract.


    """
    mail_data[key] = [  # type: ignore
        {"name": name, "smtp_address": address}
        for name, address in getaddresses(message.get_all(key, []))
    ]




def extract_body_from_eml(message: EmailMessage) -> str:
    """Extracts payload from email message.


    Args:
        message (EmailMessage): Message to extract.


    Returns:
        str: Extracted message body.
    """
    body = ""
    if message.is_multipart():
        for part in message.walk():
            main_type = part.get_content_maintype()
            content_disposition = part.get_content_disposition()
            if main_type == "text" and content_disposition != "attachment":
                charset = get_charset(part)
                body = part.get_payload(decode=True).strip().decode(charset)
    else:
        charset = get_charset(message)
        body = message.get_payload(decode=True).strip().decode(charset)
    return body


def get_charset(part: EmailMessage) -> str:
    """Gets charset from email message header.


    Args:
        part (EmailMessage): Message to extract.


    Returns:
        str: Charset from header.
    """
    try:
        return part.get_params()[1][1]
    except IndexError:
        return "utf-8"


class MailAdapter:
    """Abstract mail adapter class. Uses adapter & factory design patterns."""


    supported_extension: str = ""


    def __new__(cls, file: BinaryIO):
        """Create instance of appropriate subclass using the extensions property."""


        _, file_extension = os.path.splitext(file.name)
        if file_extension not in [".eml",".msg"]:
            raise ValueError(f"Unsupported mail extension {file_extension} given.")
        if cls is MailAdapter:
            for subclass in MailAdapter.__subclasses__():
                if not subclass.supported_extension:
                    raise RuntimeError("Supported formats property is needed")
                if file_extension == subclass.supported_extension:
                    return super().__new__(subclass)
        return super().__new__(cls)


    def __init__(self, file: BinaryIO):
        self.file = file
        self.mail_data: Email = {
            "subject": "",
            "body": "",
            "received_on": "",
            "sender": {},
            "author": {},
            "to": [],
            "cc": [],
            "attachment_names": [],
            "attachment_binaries": [],
            "attachment_types": [],
        }


    def decode(self, read_attachments: Optional[bool] = True):
        """Abstract method for decoding a file and extracting email."""
        raise NotImplementedError()




class EmlAdapter(MailAdapter):
    supported_extension = ".eml"


    def decode(self, read_attachments: Optional[bool] = True) -> Optional[Email]:
        try:
            return self._decode_from_file(self.file, read_attachments)
        except Exception:
            return None


    def _decode_from_file(
        self, file: BinaryIO, read_attachments: Optional[bool]
    ) -> Email:
        """Decodes from .eml file.


        Args:
            file (BinaryIO): File to extract.
            read_attachments (Optional[bool]): Include attachments.


        Returns:
            Email: Extracted email.
        """
        mail_data = self.mail_data
        message = message_from_binary_file(file)
        extract_metadata_from_eml_header(message, mail_data)


        mail_data["body"] = extract_body_from_eml(message)


        if read_attachments is True:
            extract_attachments_from_eml(message, mail_data)
        return mail_data
def decode_msg_body(
    body: Union[str, bytes], encodings: Tuple[str, ...] = ENCODINGS
) -> str:
    """Decodes message body based on different encodings.

    Args:
        body (Union[str, bytes]): Body to decode.
        encodings (Tuple[str, ...], optional): Encodings to try.

    Returns:
        str: Decoded body.
    """
    if isinstance(body, str):
        return body
    for encoding in encodings:
        try:
            return body.decode(encoding)
        except UnicodeDecodeError:
            pass
    return body.decode("ascii", errors="ignore")  # pragma: no cover



class MsgAdapter(MailAdapter):
    supported_extension = ".msg"

    def decode(self, read_attachments: Optional[bool] = True) -> Optional[Email]:
        try:
            return self._decode_from_file(self.file, read_attachments)
        except Exception as e:
            print(e)
            return None

    def _decode_from_file(
        self, file: BinaryIO, read_attachments: Optional[bool]
    ) -> Email:
        """Decodes from .msg file.

        Args:
            file (BinaryIO): File to extract.
            read_attachments (Optional[bool]): Include attachments.

        Returns:
            Email: Extracted email.
        """

        mail = file.read()
        mail_data = self.mail_data

        message = Message(mail)

        extract_metadata_from_eml_header(message.header, mail_data)

        if message.body is not None:
            mail_data["body"] = decode_msg_body(message.body).strip("\x00")

        if read_attachments is True:
            for attachment in message.attachments:
                name = attachment.shortFilename
                if name is not None:
                    content_type = str(guess_type(name)[0])
                    bin_data = attachment.data
                    mail_data["attachment_names"].append(name)
                    mail_data["attachment_binaries"].append(bin_data)
                    mail_data["attachment_types"].append(content_type)

        return mail_data
    
content_types = {".aac":"audio/aac",
".abw":"application/x-abiword",
".arc":"application/x-freearc",
".avif":"image/avif",
".avi":"video/x-msvideo",
".azw":"application/vnd.amazon.ebook",
".bin":"application/octet-stream",
".bmp":"image/bmp",
".bz":"application/x-bzip",
".bz2":"application/x-bzip2",
".cda":"application/x-cdf",
".csh":"application/x-csh",
".css":"text/css",
".csv":"text/csv",
".doc":"application/msword",
".docx":"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
".eot":"application/vnd.ms-fontobject",
".epub":"application/epub+zip",
".gz":"application/gzip",
".gif":"image/gif",
".htm":"text/html",
".ico":"image/vnd.microsoft.icon",
".ics":"text/calendar",
".jar":"application/java-archive",
".jpeg":"image/jpeg",
".js":"text/javascript",
".json":"application/json",
".jsonld":"application/ld+json",
".mid":"audio/midi audio/x-midi",
".mjs":"text/javascript",
".mp3":"audio/mpeg",
".mp4":"video/mp4",
".mpeg":"video/mpeg",
".mpkg":"application/vnd.apple.installer+xml",
".odp":"application/vnd.oasis.opendocument.presentation",
".ods":"application/vnd.oasis.opendocument.spreadsheet",
".odt":"application/vnd.oasis.opendocument.text",
".oga":"audio/ogg",
".ogv":"video/ogg",
".ogx":"application/ogg",
".opus":"audio/opus",
".otf":"font/otf",
".png":"image/png",
".pdf":"application/pdf",
".php":"application/x-httpd-php",
".ppt":"application/vnd.ms-powerpoint",
".pptx":"application/vnd.openxmlformats-officedocument.presentationml.presentation",
".rar":"application/vnd.rar",
".rtf":"application/rtf",
".sh":"application/x-sh",
".svg":"image/svg+xml",
".swf":"application/x-shockwave-flash",
".tar":"application/x-tar",
".tif":"image/tiff",
".ts":"video/mp2t",
".ttf":"font/ttf",
".txt":"text/plain",
".vsd":"application/vnd.visio",
".wav":"audio/wav",
".weba":"audio/webm",
".webm":"video/webm",
".webp":"image/webp",
".woff":"font/woff",
".woff2":"font/woff2",
".xhtml":"application/xhtml+xml",
".xls":"application/vnd.ms-excel",
".xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
".xml":"application/xml",
".xul":"application/vnd.mozilla.xul+xml",
".zip":"application/zip",
".3gp":"audio/video",
".3g2":"audio/video",
".7z":"application/x-7z-compressed"}


