from __future__ import print_function

import os
import sys
import time


REPOSITORY_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), os.pardir, os.pardir)
)
if REPOSITORY_ROOT not in sys.path:
    sys.path.insert(0, REPOSITORY_ROOT)

from xp.clippy_agent import ClippyController


def wait(controller, seconds):
    deadline = time.time() + seconds
    while time.time() < deadline:
        controller.pump_messages()
        time.sleep(0.05)


def main():
    controller = ClippyController()
    try:
        controller.connect()
        animations = controller.animation_names()
        print("ANIMATIONS={0}".format(",".join(animations)))
        controller.show()
        wait(controller, 2)
        controller.move(650, 420)
        wait(controller, 3)
        controller.play("Greeting")
        wait(controller, 3)
        controller.think("The persistent Python controller is talking to real Clippy.")
        print("VISIBLE_SMOKE_READY")
        wait(controller, 12)
        controller.hide()
        wait(controller, 1)
        print("VISIBLE_SMOKE_COMPLETE")
        return 0
    finally:
        controller.close()


if __name__ == "__main__":
    sys.exit(main())
