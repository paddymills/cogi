set windows-shell := ["powershell.exe", "-NoLog", "-Command"]

default:
    just --list

alias m := match
alias p := pull

# match *args:
#     python src/analysis_match_old.py {{args}}
match *args:
    python src/analysis.py --analyze {{args}}
pull *args:
    python src/analysis.py --pull {{args}}
mm *args:
    python src/analysis.py --not-matched {{args}}
analyze *args:
    python src/analysis.py {{args}}

co13:
    python src/robot/revconf.py
co02:
    python src/robot/delete.py
mbst:
    python src/mbst.py