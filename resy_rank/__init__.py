"""resy-rank — read-only Resy availability scanner and quality ranker.

This package never books, holds, modifies, or cancels a reservation. Every
Resy endpoint it touches is a GET (or the read-only venue search POST), and
there is deliberately no code path to /3/book or /3/details.
"""

__version__ = "0.1.0"
