
from argparse import ArgumentParser

from lib.db import SndbConnection

import re

def main():
    parser = ArgumentParser()
    parser.add_argument("--jobs", action="store_true", help="Print list of jobs to query")

    args = parser.parse_args()

    if args.jobs:
        get_jobs()

def get_jobs():
    pattern = re.compile(r"\d{7}[A-Za-z]$")

    with SndbConnection() as db:
        db.cursor.execute("""
            select distinct Data1 as Job
            from PartArchive
            where ArcDateTime >= '2022-01-01'
        """)
        jobs = list()
        for row in db.collect_into_namespace():
            if pattern.match(row.Job):
                jobs.append(row.Job + "*")

    with open('temp/jobs.txt', 'w') as f:
        f.write("\n".join(sorted(jobs)))


if __name__ == "__main__":
    main()
