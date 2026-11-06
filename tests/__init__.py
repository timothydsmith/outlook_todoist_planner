"""Tests for the planner. Run from the project folder with:

    py -m unittest            (Windows)
    python3 -m unittest       (Mac)

They use only the standard library, need no accounts or network, and pass in any time zone.
"""
import os

# Tests run small fake servers on 127.0.0.1. Make sure no system proxy gets in the way.
os.environ["NO_PROXY"] = "127.0.0.1,localhost"
os.environ["no_proxy"] = "127.0.0.1,localhost"
