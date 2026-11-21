"""
File-name rules shared by the launcher and the analysis scripts.

The protocol is the last '_'-separated word of the file name, after dropping the
extension and a Pulse3D '__full__<export time>' ending (same rule as the R plate
report):

    d3_SR4370_Sept26_DMD_75Hz.csv -> 75Hz
    ML25101402-2__2025_05_20_230523_d4_DAEMTR672C_75Hz__full__2025-05-23_15-02-14.xlsx -> 75Hz
"""

import os
import re

PULSE3D_SUFFIX = re.compile(r"__full__.*$")


def protocol_from_filename(filename):
    """Protocol token at the end of the file name, or None if there is no '_'."""
    stem, _ = os.path.splitext(os.path.basename(filename))
    stem = PULSE3D_SUFFIX.sub("", stem)
    if "_" not in stem:
        return None
    return stem.rsplit("_", 1)[1] or None
