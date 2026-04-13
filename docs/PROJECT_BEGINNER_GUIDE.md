# TSMMIS MATLAB Project — Beginner-Friendly Full Guide (Compact)

## 1. What this project is used for

This project is for **steel surface defect recognition** when you have:
- many images available,
- but only a few labeled examples per class (few-shot setting).

It is useful in:
- factory visual quality inspection,
- low-label industrial datasets,
- quick model setup when full annotation is expensive.

---

## 2. What model is used

Current pipeline uses:
1. **Primary model:** MATLAB built-in **ImageNet-pretrained ResNet18** (`resnet18`)  
   - Feature layer used: `pool5`  
   - Feature size: **512**
2. **Fallback model (if `resnet18` is unavailable):** project custom `ResNet18Encoder`.

### Why this model is special here
- ResNet18 is lightweight and stable.
- ImageNet pretraining gives strong general image features.
- Good fit for few-shot because support samples are very small.

---

## 3. Data split and what “unlabeled” means

Code settings:
- `trainRatio = 0.6` (60% train, 40% test)
- `labelCounts = [1,2,3,4]` for few-shot evaluation
- 6 classes (`cra, in, pa, ps, rs, sc`)

In code, “unlabeled data” means:
- training images used **without class labels** during pretraining.
- labels may exist in files, but pretraining stage does not use them.

For your dataset size (240 images/class), under current code logic:
- Train per class = `floor(240*0.6)=144`
- Reserved few-shot per class = 4
- Unlabeled per class = 140
- Test per class = 96

Expected totals:
- **Unlabeled:** 840
- **Few-shot reserved:** 24
- **Test:** 576

---

## 4. Full pipeline (Step-by-step)

## Step 1: Data preparation
- Loads images by class folder.
- Randomly splits into train/test.
- Builds:
  - `unlabeledData`
  - `fewShotData` (reserved)
  - `testData`
  - `trainPoolData` (full train pool for true N-shot sampling)

## Step 2: Model initialization
- Detects GPU (`canUseGPU`).
- Sets execution environment: `gpu` or `cpu`.
- Loads ResNet18 (or fallback).

## Step 3: “Self-supervised pretraining” in this codebase
Current implementation does:
1. Extract features from unlabeled images.
2. Normalize features.
3. PCA projection (`pcaDim = 256`).
4. K-means initialization (`numClasses = 6`).
5. **300 epochs** of prototype-refinement updates.

Important note:
- In this current code, `epochs` are actively used.
- `learningRate/momentum/weightDecay` are present in config but are **not directly used** inside this specific pretraining function.

## Step 4: Few-shot evaluation
For each label count in `[1,2,3,4]`:
1. Run **100 trials** (`numRuns = 100`).
2. In each trial:
   - sample `N` labels per class from full training pool,
   - train prototype classifier from support features,
   - adapt prototypes using unlabeled features (`selfTrainIterations=2`, margin `0.08`),
   - predict on test set.
3. Save all trial accuracies.

## Step 5: Results table
- Reports per setting:
  - **Mean Accuracy**
  - **Std Dev**

## Step 6: Feature visualization
- Builds 2D embedding using `tsne` (fallback to PCA if needed).
- Saves embedding and plot safely (no crash if GUI export fails).

---

## 5. What “epochs”, “labels_1..4”, mean accuracy, std dev mean

### Epochs (in this code)
- 1 epoch = one full update cycle of class prototypes in pretraining.
- Total pretraining epochs configured: **300**.

### `labels_1`, `labels_2`, `labels_3`, `labels_4`
These are **not class IDs**.  
They mean:
- `labels_1`: 1 labeled sample per class
- `labels_2`: 2 labeled samples per class
- `labels_3`: 3 labeled samples per class
- `labels_4`: 4 labeled samples per class

### Mean Accuracy
- Average of 100 trial accuracies for a label setting.

### Std Dev
- Standard deviation of those 100 trial accuracies.
- Lower std dev = more stable performance across random few-shot samples.

---

## 6. Your real reported results (from latest run)

| Setting | Meaning | Mean Accuracy | Std Dev |
|---|---|---:|---:|
| `labels_1` | 1 label/class | 88.13% | 4.68% |
| `labels_2` | 2 labels/class | 93.11% | 2.64% |
| `labels_3` | 3 labels/class | 94.86% | 1.85% |
| `labels_4` | 4 labels/class | 95.75% | 1.51% |

Interpretation:
- Accuracy increases as labeled support per class increases.
- Stability also improves (std dev decreases).

---

## 7. What changed that improved accuracy (short)

Main accuracy gains came from:
1. **True N-shot sampling from full train pool** (not fixed tiny support).
2. **Real feature-pretraining flow** (feature extraction + PCA + prototype refinement) replacing placeholder behavior.
3. **Prototype classifier + unlabeled adaptation** for few-shot robustness.
4. **Combined prototype and support-similarity scoring** at prediction time.

Also fixed:
- robust t-SNE saving (no crash at visualization step).
- GPU path for feature extraction.

---

## 8. Result files and how they can be used later

Saved outputs:
- `models/pretrained_encoder.mat`
  - contains `net`, `trainInfo`, `modelInfo`
- `evaluation/results/evaluation_results.mat`
  - contains `results`, `config`, `modelInfo`, `trainInfo`
- `evaluation/results/tsne_embedding.mat`
  - contains 2D embedding + labels
- `evaluation/results/tsne_visualization.png`
  - saved plot image (if export available)

### Can these be used further?
Yes:
- compare experiments/hyperparameters over time,
- make reports/plots without rerunning all experiments,
- reuse saved model/features for additional testing,
- benchmark on other defect datasets with the same pipeline.

---

## 9. Practical limitations (important)

- This pipeline is **feature-based** in current main script (not full end-to-end gradient fine-tuning in Step 4).
- Performance depends on data quality and class balance.
- Few-shot random sampling introduces variance (that is why std dev is reported).

---

## 10. One-line summary

This project is a practical few-shot defect recognition pipeline using ResNet18 features + unlabeled adaptation; after fixing data sampling and evaluation logic, your measured accuracy improved strongly (88% to 95% across 1–4 shot settings).

