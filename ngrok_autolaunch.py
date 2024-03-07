import dataclasses
import json

from pyngrok import ngrok
from pyngrok.ngrok import NgrokTunnel


@dataclasses.dataclass
class tunnel:
    public_url: str
    localport: int

    @staticmethod
    def from_(tun: NgrokTunnel):
        public_url = tun.public_url
        localport = int(tun.config['addr'].split(':')[-1])
        return tunnel(public_url=public_url, localport=localport)


# Open a HTTP tunnel on the default port 80
# <NgrokTunnel: "https://<public_sub>.ngrok.io" -> "http://localhost:80">
try:
    _ = ngrok.connect(name='cat')
    print(_)
    cat_tunnel = tunnel.from_(_)

    _ = ngrok.connect(name='nca')
    print(_)
    nca_tunnel = tunnel.from_(_)

    _ = ngrok.connect(name='goo')
    print(_)
    goo_tunnel = tunnel.from_(_)

    with open('secrets/tunnels', 'w') as f:
        json.dump([dataclasses.asdict(t) for t in [cat_tunnel, nca_tunnel, goo_tunnel]], f)
except Exception:
    with open('secrets/tunnels', 'r') as f:
        params = json.load(f)
        cat_tunnel, nca_tunnel, goo_tunnel = map(lambda d: tunnel(**d), params)
