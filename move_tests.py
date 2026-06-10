import re

with open("src/app/backstage/mod.rs", "r") as f:
    content = f.read()

start_str = "#[cfg(test)]\nmod tests {"
start_idx = content.find(start_str)
if start_idx == -1:
    print("Tests not found!")
    exit(1)

# Find the matching closing brace for the module
brace_count = 0
end_idx = -1
in_tests = False
for i in range(start_idx, len(content)):
    if content[i] == '{':
        brace_count += 1
        in_tests = True
    elif content[i] == '}':
        brace_count -= 1
        if in_tests and brace_count == 0:
            end_idx = i + 1
            break

if end_idx == -1:
    print("Could not find end of tests module!")
    exit(1)

tests_content = content[start_idx:end_idx]
new_content = content[:start_idx] + content[end_idx:] + "\n" + tests_content + "\n"

with open("src/app/backstage/mod.rs", "w") as f:
    f.write(new_content)

print("Moved tests.")
