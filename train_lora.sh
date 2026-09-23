#!/bin/bash
# Train the teacher-voice LoRA adapter from the current labeled corpus.
#
# Rebuilds the dataset from the labeling sheet, runs MLX LoRA fine-tuning,
# then copies the checkpoint with the LOWEST validation loss over the final
# weights -- with a corpus this small the model memorizes the training set
# within ~100 iterations, so early stopping by val loss is what actually
# picks a usable adapter. Whenever new corpus arrives, just re-run this
# script: a full retrain on old+new together beats resuming from the old
# adapter (no forgetting, no batch-order bias) and only takes minutes.
set -euo pipefail
cd "$(dirname "$0")"

MODEL="mlx-community/Qwen2.5-7B-Instruct-4bit"
DATA="app_data/training"
ADAPTERS="$DATA/adapters"
LOG="$ADAPTERS/train.log"

.venv/bin/python build_training_set.py
echo
mkdir -p "$ADAPTERS"
rm -f "$ADAPTERS"/0*_adapters.safetensors
.venv/bin/python -m mlx_lm lora \
  --model "$MODEL" \
  --train \
  --data "$DATA" \
  --fine-tune-type lora \
  --num-layers 8 \
  --batch-size 1 \
  --iters 400 \
  --learning-rate 5e-5 \
  --steps-per-eval 25 \
  --val-batches -1 \
  --save-every 25 \
  --max-seq-length 1024 \
  --grad-checkpoint \
  --mask-prompt \
  --adapter-path "$ADAPTERS" 2>&1 | tee "$LOG"

.venv/bin/python - "$ADAPTERS" "$LOG" <<'PY'
import re, shutil, sys
from pathlib import Path

adapters, log = Path(sys.argv[1]), Path(sys.argv[2])
vals = {int(m[1]): float(m[2])
        for m in re.finditer(r"Iter (\d+): Val loss ([\d.]+)", log.read_text())}
saved = {int(p.name.split("_")[0]): p for p in adapters.glob("0*_adapters.safetensors")}
scored = {it: vals[it] for it in saved if it in vals}
if not scored:
    sys.exit("no checkpoint has a recorded val loss; keeping final weights")
best = min(scored, key=scored.get)
shutil.copyfile(saved[best], adapters / "adapters.safetensors")
print(f"best checkpoint: iter {best} (val loss {scored[best]:.3f}) -> adapters.safetensors")
PY
