
SRC_DIR := src

# catch-all for launching a binary not explicity listed
%: $(SRC_DIR)/%.py
	python $^

rollup:
	python src/rollup.py -c

# move .ready files
mv: $(SRC_DIR)/*.ready
	python src/utils.py --move $^
