import sys, re
sys.stdout.reconfigure(encoding='utf-8')
text = open(r'e:\Working Documents\【Important】My Projects\24 Specific avatar\interrogation_app\templates\manage.html', encoding='utf-8').read()
m = re.search(r'(<div class="admin-modal-overlay" id="rescheduleModal".*?</div>\n</div>)', text, re.DOTALL)
print(m.group(1) if m else 'Not found')
