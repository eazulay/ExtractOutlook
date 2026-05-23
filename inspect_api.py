import pypff
import re
from email.utils import parsedate_to_datetime

f = pypff.file()
f.open(r"c:\node\ExtractOutlook\Archive GIW to 2004.pst")
root = f.get_root_folder()

# Walk all messages and find ones with empty delivery_time
def find_empty_delivery(folder, path="", found=[]):
    name = folder.get_name() or "unnamed"
    current = path + "/" + name
    for i in range(folder.get_number_of_sub_messages()):
        msg = folder.get_sub_message(i)
        try:
            dt = msg.get_delivery_time()
        except Exception:
            dt = None
        if not dt:
            found.append((current, msg))
        if len(found) >= 3:
            return
    for i in range(folder.get_number_of_sub_folders()):
        if len(found) >= 3:
            return
        find_empty_delivery(folder.get_sub_folder(i), current, found)

found = []
for i in range(root.get_number_of_sub_folders()):
    find_empty_delivery(root.get_sub_folder(i), "", found)
    if len(found) >= 3:
        break

print(f"Found {len(found)} messages with empty delivery_time\n")

date_methods = [m for m in dir(pypff.message) if "time" in m.lower() or "date" in m.lower() or "submit" in m.lower() or "creat" in m.lower() or "modif" in m.lower()]
print("Date-related methods on message:", date_methods, "\n")

for folder_path, msg in found:
    print(f"--- {folder_path} | subject: {msg.get_subject()!r}")
    for method_name in date_methods:
        try:
            val = getattr(msg, method_name)()
            print(f"  {method_name}() = {val!r}")
        except Exception as e:
            print(f"  {method_name}() -> ERROR: {e}")

    headers = None
    try:
        headers = msg.get_transport_headers()
    except Exception:
        pass

    if headers:
        m = re.search(r"^Date:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
        if m:
            raw_date = m.group(1).strip()
            print(f"  Date header: {raw_date!r}")
            try:
                parsed = parsedate_to_datetime(raw_date)
                print(f"  Parsed Date header: {parsed.isoformat()}")
            except Exception as e:
                print(f"  Parse error: {e}")
        else:
            print("  No Date: header found")
        print(f"  Headers snippet: {headers[:300]!r}")
    else:
        print("  No transport headers")
    print()

f.close()
