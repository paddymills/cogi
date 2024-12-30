set windows-shell := ["powershell.exe", "-NoLog", "-Command"]

default:
    just --list

alias m := match
alias p := pull

# match *args:
#     python src/analysis_match_old.py {{args}}
match *args:
    uv run analysis --analyze {{args}}
pull *args:
    uv run analysis --pull {{args}}
mm *args:
    uv run analysis --not-matched {{args}}
analyze *args:
    uv run analysis {{args}}

co13:
    uv run revconf
co02:
    uv run delete
mbst:
    uv run mbst