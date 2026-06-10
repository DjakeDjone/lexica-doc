import re

with open("src/app/mod.rs", "r") as f:
    content = f.read()

# Find the handle_api_request function
start_str = "#[cfg(not(target_arch = \"wasm32\"))]\npub fn handle_api_request("
start_idx = content.find(start_str)
if start_idx == -1:
    print("Function not found!")
    exit(1)

end_idx = content.find("\n}\n", start_idx) + 3

func_content = content[start_idx:end_idx]

# Remove from app/mod.rs
new_app_content = content[:start_idx] + content[end_idx:]
# Replace handle_api_request with crate::http_server::handle_api_request
new_app_content = new_app_content.replace("handle_api_request(\n                    request,", "crate::http_server::handle_api_request(\n                    request,")

with open("src/app/mod.rs", "w") as f:
    f.write(new_app_content)

# Append to http_server.rs
with open("src/http_server.rs", "a") as f:
    f.write("\n" + func_content)

print("Moved function.")
