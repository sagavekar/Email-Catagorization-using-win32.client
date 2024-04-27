from notifypy import Notify

notification = Notify()
notification.application_name = "test"
notification.title = "test"
notification.message = "test"
notification.send()