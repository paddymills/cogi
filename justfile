set windows-shell := ["powershell.exe", "-NoLog", "-Command"]

default:
    just --list

alias m := match
alias p := pull

match *args:
    python src/analysis_match_old.py {{args}}
match_dev *args:
    python src/analysis_match.py {{args}}
pull *args:
    python src/analysis_pull.py {{args}}

co13:
    python src/robot/revconf.py
co02:
    python src/robot/delete.py