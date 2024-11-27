import subprocess
# Opening Hotspot
def run_netsh_command(command):
    try:
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        print(e.stdout.decode('utf-8'))
        print(e.stderr.decode('utf-8'))
        return False
    return True
# Configure the Wi-Fi adapter to allow hosted networks
enable_hosted_network = ['netsh', 'wlan', 'set', 'hostednetwork', 'mode=allow', 'ssid=PrintVendo', 'key=12345678']

# Start the hosted network
start_hosted_network = ['netsh', 'wlan', 'start', 'hostednetwork']

# Run the commands
if run_netsh_command(enable_hosted_network):
    print("Hosted network enabled successfully.")
if run_netsh_command(start_hosted_network):
    print("Hosted network started successfully.")
