"""Builds the Power Automate flow definition (Dataverse clientdata) for the Kovie relay.

Apps Script emails [removed] a relay message whose body carries
KOVIEB64:<base64 JSON>:KOVIEEND. The flow decodes it, checks the shared secret and
the recipient domain, then sends the email (reply-to Kate) and creates or deletes
the Outlook calendar invite from Reese's UARK account.

Usage: python3 build_flow.py <secret> <connectionName> <relaySenderGmail>  -> prints clientdata JSON
The relay sender (the Google account running Apps Script) is also allowed as a test recipient.
"""
import json, sys

secret, conn, sender = sys.argv[1], sys.argv[2], sys.argv[3].lower()
P = "outputs('Payload')"
api = {"connectionName": "shared_office365", "operationId": None, "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365"}

def op(operation, params, after=None):
    return {
        "type": "OpenApiConnection",
        "inputs": {
            "host": dict(api, operationId=operation),
            "parameters": params,
            "authentication": "@parameters('$authentication')",
        },
        "runAfter": after or {},
    }

payload_expr = ("@json(base64ToString(first(split(last(split(triggerOutputs()?['body/body'], 'KOVIEB64:')), ':KOVIEEND'))))")

calendar_id = "@first(body('Filter_calendars'))?['id']"

definition = {
    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
        "$connections": {"defaultValue": {}, "type": "Object"},
        "$authentication": {"defaultValue": {}, "type": "SecureObject"},
    },
    "triggers": {
        "When_a_relay_email_arrives": {
            "type": "OpenApiConnectionNotification",
            "splitOn": "@triggerOutputs()?['body/value']",
            "inputs": {
                "host": dict(api, operationId="OnNewEmailV3"),
                "parameters": {
                    "folderPath": "Inbox",
                    "from": sender,
                    "subjectFilter": "[KOVIE-RELAY]",
                    "importance": "Any",
                    "fetchOnlyWithAttachment": False,
                    "includeAttachments": False,
                },
                "authentication": "@parameters('$authentication')",
            },
        }
    },
    "actions": {
        "Payload": {"type": "Compose", "inputs": payload_expr, "runAfter": {}},
        "Is_valid_relay": {
            "type": "If",
            "runAfter": {"Payload": ["Succeeded"]},
            "expression": {"and": [
                {"equals": [f"@{P}?['secret']", secret]},
                {"or": [
                    {"endsWith": [f"@toLower({P}?['to'])", "uark.edu"]},
                    {"equals": [f"@toLower({P}?['to'])", sender]},
                ]},
            ]},
            "actions": {
                "Send_email": op("SendEmailV2", {
                    "emailMessage/To": f"@{P}?['to']",
                    "emailMessage/Subject": f"@{P}?['subject']",
                    "emailMessage/Body": f"<p>@{{{P}?['html']}}</p>",
                    "emailMessage/ReplyTo": f"@{P}?['replyTo']",
                    "emailMessage/Importance": "Normal",
                }),
                "Get_calendars": op("CalendarGetTables_V2", {}, {"Send_email": ["Succeeded"]}),
                "Filter_calendars": {
                    "type": "Query",
                    "runAfter": {"Get_calendars": ["Succeeded"]},
                    "inputs": {"from": "@body('Get_calendars')?['value']", "where": "@equals(item()?['name'], 'Calendar')"},
                },
                "Has_invite": {
                    "type": "If",
                    "runAfter": {"Filter_calendars": ["Succeeded"]},
                    "expression": {"and": [{"not": {"equals": [f"@{P}?['invite']", "@null"]}}]},
                    "actions": {
                        "Create_invite": op("V4CalendarPostItem", {
                            "table": calendar_id,
                            "item/subject": f"@{P}?['invite']?['subject']",
                            "item/start": f"@{P}?['invite']?['start']",
                            "item/end": f"@{P}?['invite']?['end']",
                            "item/timeZone": "(UTC-06:00) Central Time (US & Canada)",
                            "item/requiredAttendees": f"@{P}?['to']",
                            "item/body": f"<p>@{{{P}?['invite']?['body']}}</p>",
                            "item/location": f"@{P}?['invite']?['location']",
                            "item/importance": "normal",
                            "item/reminderMinutesBeforeStart": 60,
                            "item/isReminderOn": True,
                            "item/showAs": "busy",
                            "item/responseRequested": True,
                        }),
                    },
                    "else": {"actions": {}},
                },
                "Has_uninvite": {
                    "type": "If",
                    "runAfter": {"Has_invite": ["Succeeded"]},
                    "expression": {"and": [{"not": {"equals": [f"@{P}?['uninvite']", "@null"]}}]},
                    "actions": {
                        "Get_slot_events": op("GetEventsCalendarViewV3", {
                            "calendarId": calendar_id,
                            "startDateTimeUtc": f"@{P}?['uninvite']?['startUtc']",
                            "endDateTimeUtc": f"@{P}?['uninvite']?['endUtc']",
                        }),
                        "Filter_actor_events": {
                            "type": "Query",
                            "runAfter": {"Get_slot_events": ["Succeeded"]},
                            "inputs": {
                                "from": "@body('Get_slot_events')?['value']",
                                "where": f"@and(startsWith(item()?['subject'], 'Kovie actor slot'), contains(string(item()), {P}?['uninvite']?['key']))",
                            },
                        },
                        "Delete_each": {
                            "type": "Foreach",
                            "foreach": "@body('Filter_actor_events')",
                            "runAfter": {"Filter_actor_events": ["Succeeded"]},
                            "actions": {
                                "Delete_invite": op("CalendarDeleteItem_V2", {"calendar": calendar_id, "event": "@items('Delete_each')?['id']"}),
                            },
                        },
                    },
                    "else": {"actions": {}},
                },
                "Delete_relay_email": op("DeleteEmail_V2", {"messageId": "@triggerOutputs()?['body/id']"}, {"Has_uninvite": ["Succeeded"]}),
            },
            "else": {"actions": {}},
        },
    },
    "outputs": {},
}

clientdata = {
    "properties": {
        "connectionReferences": {
            "shared_office365": {
                "runtimeSource": "embedded",
                "connection": {"name": conn},
                "api": {"name": "shared_office365"},
            }
        },
        "definition": definition,
    },
    "schemaVersion": "1.0.0.0",
}
print(json.dumps(clientdata))
