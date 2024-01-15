
from argparse import ArgumentParser
from datetime import datetime
from glob import glob
from tqdm import tqdm

import os


def main():
    parser = ArgumentParser()
    parser.add_argument("-p", "--program", help="program to issue")

    args = parser.parse_args()

    if args.program:
        issue_file_from_prod(get_program_line(args.program))


def get_program_line(program):
    found = False
    # path = os.path.join(os.environ['USERPROFILE'], r"Documents\sapcnf\outbound\Production_*.outbound.archive")
    path = os.path.join(os.environ['USERPROFILE'], r"\\hiifileserv1\sigmanestprd\Archive\Production_*.outbound.archive")
    for fn in tqdm(sorted(glob(path), reverse=True), desc="parsing files"):
        with open(fn) as f:
            for line in f.readlines():
                s = line.strip().split('\t')
                if len(s) > 10:
                    prog = s[12]

                    if prog == program:
                        found = True
                        yield s

        if found:
            return


def issue_file_from_prod(lines):
    qty = 0.0
    line = None
    for row in lines:
        line = row
        print(row)
        qty += float(row[8])
    line[8] = f"{qty:.3f}"

    if not line[4]:
        code = "PR02"
    elif line[7][2:9] == line[2][2:9]:
        code = "PR01"
    else:
        code = "PR03"
    issue = [code, line[1].replace('S', 'D'), '01', *line[6:]]
    fn = 'Issue_{}.ready'.format(datetime.now().strftime("%Y%m%d%H%M%S"))
    with open(os.path.join(os.path.dirname(__file__), fn), 'w') as f:
        f.write('\t'.join(issue))


if __name__ == "__main__":
    main()
