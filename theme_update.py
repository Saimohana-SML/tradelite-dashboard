import glob
import re

print("Starting theme update...")
for f in glob.glob('e:/tradelite/*.html'):
    with open(f, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Replace tailwind blue classes with orange
    content = re.sub(r'(-|:)blue-(\d{2,3})', r'\1orange-\2', content)
    
    # Replace specific hex colors in Tailwind config blocks (for employees & admin-dashboard)
    content = content.replace('"#003d9b"', '"#ea580c"') # primary orange-600
    content = content.replace('"#0052cc"', '"#f97316"') # primary-container orange-500
    content = content.replace('"#dae2ff"', '"#ffedd5"') # secondary-fixed orange-100
    content = content.replace('"#021945"', '"#7c2d12"') # on-secondary-fixed orange-900
    
    with open(f, 'w', encoding='utf-8') as file:
        file.write(content)
    print(f"Updated {f}")
print("Theme color updated successfully!")
