# TSMMIS MATLAB Project: Full Info + Accuracy Improvement Change Log

## 1) Project Summary

This project implements **TSMMIS-style self-supervised few-shot steel defect recognition** on the NEU-CLS dataset (6 classes: `cra`, `in`, `pa`, `ps`, `rs`, `sc`).

Main goals:
- Learn strong visual representations from mostly unlabeled data.
- Evaluate few-shot recognition with 1/2/3/4 labels per class.
- Support GPU acceleration when CUDA is available.

Latest real results from your run:
- **1-shot:** 88.13 +/- 4.68%
- **2-shot:** 93.11 +/- 2.64%
- **3-shot:** 94.86 +/- 1.85%
- **4-shot:** 95.75 +/- 1.51%

---

## 2) All Accuracy-Related Changes Made

## A. `utils/DataLoader.m`

### Change
- `loadDataset(...)` now returns a **4th output**: `trainPoolData` (full training pool for each class).

### Why this improved accuracy
- Previously, evaluation for label counts 2/3/4 still depended on a tiny reserved set and could not properly sample true N-shot episodes.
- Now, each run can sample from the full train pool per class, so 1/2/3/4-shot settings are real and meaningful.

### Exact behavior difference
- Before: few-shot set was effectively fixed/small for all runs.
- After: few-shot support is sampled from full train pool each run, per class.

---

## B. `main_TSMMIS.m` (rewritten pipeline core)

### Change 1: Replaced placeholder training behavior with real feature pretraining flow
- Added `initializeEncoder(config)`:
  - Uses pretrained `resnet18` if available.
  - Falls back to project encoder otherwise.
- Added `pretrainModel(net, unlabeledData, opts)`:
  - Extracts real features from unlabeled data.
  - Performs PCA projection + prototype refinement (iterative clustering-style update).
  - Stores transform/prototype state in `net.transform`.

### Why this improved accuracy
- Original pipeline used placeholder/random-like loss behavior and did not produce useful representation adaptation.
- New flow uses actual image features and structured unlabeled pretraining.

---

### Change 2: Fixed few-shot evaluation to use full training pool + repeated randomized episodes
- Added precomputed feature banks:
  - `poolFeaturesRaw`, `unlabeledFeaturesRaw`, `testFeaturesRaw`
  - transformed via `applyFeatureTransform(...)`
- Added label extraction and true N-shot sampling:
  - `getDataLabels(...)`
  - `sampleFewShotIndices(...)`

### Why this improved accuracy
- Ensures each run truly uses N labels/class and captures episode variability.
- Avoids the previous degenerate behavior that caused chance-level flat results.

---

### Change 3: Replaced weak linear setup with robust few-shot classifier + adaptation
- Added classifier functions:
  - `trainPrototypeClassifier(...)`
  - `adaptPrototypeClassifier(...)`
  - `predictPrototypeClassifier(...)`
- Added mixed decision logic:
  - Prototype similarity + support-instance similarity fusion.
- Added unlabeled adaptation of class prototypes (self-training style refinement).

### Why this improved accuracy
- Better suited for tiny support sets than plain linear fitting on minimal labels.
- Uses unlabeled training features to stabilize class centers.

---

### Change 4: GPU acceleration wiring
- Auto-detects GPU with `canUseGPU`.
- Captures device with `gpuDevice`.
- Sets `config.executionEnvironment` (`'gpu'`/`'cpu'`).
- Uses GPU execution path for feature extraction via:
  - `activations(..., 'ExecutionEnvironment', config.executionEnvironment)`
  - GPU path in fallback feature extraction where applicable.

### Why this helps
- Faster feature extraction and overall runtime on CUDA-capable systems.
- Keeps behavior consistent (same pipeline, faster backend).

---

### Change 5: Fixed Step 6 visualization crash (graphics handshaking / invalid object)
- Replaced direct `saveas(gcf, ...)` flow with robust `saveTsneFigure(...)`:
  - Off-screen figure creation (`Visible`, `off`)
  - `exportgraphics` fallback to `saveas`
  - non-fatal handling when figure windows are unavailable
- Always saves embedding data to:
  - `evaluation/results/tsne_embedding.mat`

### Why this matters
- Prevents pipeline termination at visualization stage.
- Keeps training/evaluation results saved even in headless or unstable graphics sessions.

---

## C. `evaluation/TSMMISEvaluator.m`

### Change
- Hardened visualization export in helper functions:
  - Off-screen figures for `tsneVisualization(...)` and `plotTrainingProgress(...)`
  - safe export helper (`safeExportFigure(...)`) using `exportgraphics` or `saveas`
  - warning on export failure instead of hard crash

### Why this matters
- Prevents plotting failures from killing evaluation scripts.
- Improves reliability for batch/headless runs.

---

## D. `docs/README.md`

### Change
- Updated docs to reflect:
  - `loadDataset` 4th output (`trainPool`)
  - current pipeline calls and GPU auto-detection behavior

### Why this matters
- Documentation now matches actual executable code path.

---

## 3) Full Project Information

## Project Structure

```
TSMMIS_MATLAB/
├── main_TSMMIS.m                  % Main end-to-end pipeline (current primary entry)
├── data/NEU-CLS/                  % Dataset (6 classes)
├── models/                        % Saved models (.mat)
├── training/                      % Losses + training utilities
├── evaluation/                    % Evaluation/fine-tuning utilities + plotting helpers
├── utils/                         % Data loading and augmentations
└── docs/                          % Documentation
```

---

## Core Modules

### `main_TSMMIS.m`
- Orchestrates complete flow:
  1. Data preparation
  2. Encoder initialization
  3. Unlabeled pretraining
  4. Few-shot evaluation loop
  5. Result table
  6. t-SNE + export

### `utils/DataLoader.m`
- Loads dataset from class folders.
- Splits train/test.
- Provides:
  - unlabeled training set
  - reserved few-shot set
  - test set
  - full training pool (`trainPoolData`) for N-shot sampling.

### `models/ResNet18Encoder.m`
- Project encoder architecture (fallback path when built-in `resnet18` unavailable).

### `training/*.m`
- Includes contrastive/MMIS loss classes and trainer utilities.
- Main pipeline now uses the improved feature-pretraining path in `main_TSMMIS.m`.

### `evaluation/*.m`
- Evaluation utilities and plotting support.
- Visualization export made crash-resistant.

---

## Data Assumptions

- Dataset path: `data/NEU-CLS`
- Class folders expected directly under dataset root.
- Labels are encoded as class indices and converted to 1..numClasses for classifier internals.

---

## Runtime / Toolboxes

Required MATLAB toolboxes:
- Deep Learning Toolbox
- Image Processing Toolbox
- Statistics and Machine Learning Toolbox

Recommended:
- Parallel Computing Toolbox (for broader GPU workflows)

GPU:
- CUDA-capable GPU auto-detected by `canUseGPU`.
- Uses GPU execution for feature extraction when available.

---

## Main Configuration Knobs (`main_TSMMIS.m`)

- Data:
  - `config.dataPath`
  - `config.trainRatio`
  - `config.imageSize`
- Pretraining:
  - `config.pretrain.epochs`
  - `config.pretrain.batchSize`
  - `config.pretrain.learningRate`
  - `config.pretrain.pcaDim`
- Few-shot evaluation:
  - `config.finetune.labelCounts`
  - `config.finetune.numRuns`
  - `config.finetune.selfTrainIterations`
  - `config.finetune.selfTrainMargin`

---

## Output Artifacts

- `models/pretrained_encoder.mat`
  - Saved encoder/model state and pretraining info.
- `evaluation/results/evaluation_results.mat`
  - Mean/std/all-run accuracies for each label count.
- `evaluation/results/tsne_embedding.mat`
  - Saved 2D embedding + labels.
- `evaluation/results/tsne_visualization.png`
  - Saved t-SNE figure when graphics export is available.

---

## How to Run

From MATLAB:

```matlab
cd D:/MVRP_Project/TSMMIS_MATLAB
main_TSMMIS
```

---

## 4) Root Cause of Original Low Accuracy + Fix Summary

Original issues:
1. Placeholder/non-updating pretraining behavior.
2. Few-shot sampling did not truly scale with label count (2/3/4-shot path ineffective).
3. Plot export crash interrupted execution.

Fixes:
1. Real feature-pretraining + transform state.
2. True N-shot sampling from full train pool + improved few-shot classifier/adaptation.
3. Robust plotting export path and safe handling.

Result:
- Accuracy moved from chance-level behavior to strong few-shot performance (88%+ even at 1-shot in your run).

