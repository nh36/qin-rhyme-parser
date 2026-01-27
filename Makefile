.PHONY: test

test:
	python3 main.py --mode both --outdir outputs --pdf hudie2023_qin_rhymes.pdf && \
	python3 scripts/regression_canaries.py --manifest outputs/latest_manifest.json
