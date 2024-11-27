import cherrypy
from PrintVendo import app  # Import your Flask app instance
import socket

if __name__ == '__main__':
    # Get the local IP address dynamically
    host_ip = socket.gethostbyname(socket.gethostname())
    
    cherrypy.tree.graft(app, "/")
    cherrypy.config.update({
        'server.socket_host': host_ip,
        'server.socket_port': 80,
    })

    cherrypy.engine.start()
    cherrypy.engine.block()
