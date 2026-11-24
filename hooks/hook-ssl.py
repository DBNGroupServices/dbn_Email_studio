"""Include adjacent trust-policy data required by the bundled Python SSL module.

The Codex runtime loads this file in SSLContext.__new__. Preserve that policy
when freezing the runtime, rather than overriding or disabling verification.
Standard Python installations that don't need the file require no extra data.
"""
from pathlib import Path
import ssl

policy_file = Path(ssl.__file__).with_name('codex-rejected-cas.pem')
datas = [(str(policy_file), '.')] if policy_file.is_file() else []
