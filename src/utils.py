
import re
import shutil

from argparse import ArgumentParser
from pprint import pprint

def main():
    parser = ArgumentParser()
    parser.add_argument("--move", action="store", help="Move Production_*.ready files")
    parser.add_argument("--jobs", action="store_true", help="output jobs from `temp/parts.txt`")
    args = parser.parse_args()

    if args.move:
        shutil.move(args.move, r"\\hiifileserv1\sigmanestprd\Outbound")
    
    if args.jobs:
        get_jobs()


def get_jobs():
    with open("temp/parts.txt") as f:
        pattern = re.compile(r"\d{7}[A-Z]$")
        jobs = set([x.split("-", 1)[0].upper() for x in f.readlines()])
        for job in sorted(jobs):
            if pattern.match(job):
                print(job + "*")


if __name__ == "__main__":
    main()
