"""MCP tool implementations grouped by Outlook area."""

from typing import Any

_API_DISPATCH: dict[str, Any] | None = None


def get_api_dispatch() -> dict[str, Any]:
    global _API_DISPATCH
    if _API_DISPATCH is not None:
        return _API_DISPATCH
    from . import mail as mail_tools
    from . import calendar as cal_tools
    from . import contacts as contact_tools
    from . import tasks as task_tools
    from . import account as account_tools

    _API_DISPATCH = {
        "mail/list": mail_tools.mail_list,
        "mail/search": mail_tools.mail_search,
        "mail/get": mail_tools.mail_get,
        "mail/send": mail_tools.mail_send,
        "mail/reply": mail_tools.mail_reply,
        "mail/move": mail_tools.mail_move,
        "mail/delete": mail_tools.mail_delete,
        "mail/forward": mail_tools.mail_forward,
        "mail/mark_read": mail_tools.mail_mark_read,
        "mail/mark_unread": mail_tools.mail_mark_unread,
        "mail/attachments": mail_tools.mail_attachments,
        "mail/folders": mail_tools.mail_folders,
        "calendar/list": cal_tools.calendar_list,
        "calendar/get": cal_tools.calendar_get,
        "calendar/create": cal_tools.calendar_create,
        "calendar/update": cal_tools.calendar_update,
        "calendar/delete": cal_tools.calendar_delete,
        "calendar/search": cal_tools.calendar_search,
        "contact/list": contact_tools.contacts_list,
        "contact/search": contact_tools.contacts_search,
        "contact/get": contact_tools.contact_get,
        "contact/create": contact_tools.contact_create,
        "contact/update": contact_tools.contact_update,
        "contact/delete": contact_tools.contact_delete,
        "task/list": task_tools.tasks_list,
        "task/get": task_tools.task_get,
        "task/create": task_tools.task_create,
        "task/update": task_tools.task_update,
        "task/delete": task_tools.task_delete,
        "task/search": task_tools.task_search,
        "account/info": account_tools.account_info,
    }
    return _API_DISPATCH
