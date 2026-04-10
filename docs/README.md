# TSMMIS Implementation Guide for MATLAB

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Project Structure](#project-structure)
4. [Dataset Preparation](#dataset-preparation)
5. [Quick Start Guide](#quick-start-guide)
6. [Detailed Module Description](#detailed-module-description)
7. [Training Configuration](#training-configuration)
8. [Running the Pipeline](#running-the-pipeline)
9. [Understanding the Results](#understanding-the-results)
10. [Troubleshooting](#troubleshooting)
11. [Ablation Experiments](#ablation-experiments)

---

## 1. Introduction

This implementation provides a complete pipeline for **Self-Supervised Few-Shot Steel Surface Defect Recognition** based on the TSMMIS framework. The system uses a single-encoder architecture with multi-view contrastive learning and Multi-Modal Instance Similarity (MMIS) regularization.

**Key Features:**
- ResNet18 backbone with predictor MLP
- Multicrop augmentation (K=5 views)
- Contrastive loss (L_t&s) + MMIS regularization (L_MMIS)
- Few-shot fine-tuning with 1-4 labeled samples per class
- Linear evaluation and full fine-tuning modes

---

## 2. System Requirements

### Software Requirements
- **MATLAB R2021b or later** (recommended R2022b+)
- **Deep Learning Toolbox**
- **Image Processing Toolbox**
- **Statistics and Machine Learning Toolbox**
- **Parallel Computing Toolbox** (optional, for GPU training)

### Hardware Requirements
- **Minimum:** 8GB RAM, Intel Core i5 CPU
- **Recommended:** 16GB+ RAM, NVIDIA GPU with CUDA (for faster training)

### MATLAB Add-ons Required
1. Deep Learning Toolbox
2. Image Processing Toolbox  
3. Statistics and Machine Learning Toolbox

To verify installed add-ons:
```matlab
matlab.addons.installedAddons
```

---

## 3. Project Structure

```
TSMMIS_MATLAB/
├── main_TSMMIS.m           % Main execution script
├── data/                   % Dataset folder
│   └── NEU-CLS/           % NEU-CLS dataset (6 classes)
│       ├── cra/           % Cracling
│       ├── in/            //nclusion
│       ├── pa/            //atches
│       ├── ps/            //its
│       ├── rs/            //olls
│       └── sc/            //Cratches
├── models/
│   └── ResNet18Encoder.m  % ResNet18 architecture
├── utils/
│   └── DataLoader.m       % Data loading & multicrop
├── training/
│   ├── ContrastiveLoss.m   % L_t&s loss
│   ├── MMISLoss.m         % MMIS regularization
│   └── TSMMISTrainer.m    % Training loop
├── evaluation/
│   └── FineTuner.m        % Fine-tuning & evaluation
└── docs/
    └── README.md          % This file
```

---

## 4. Dataset Preparation

### NEU-CLS Dataset
The recommended dataset is **NEU-CLS** (6 classes, 300 images each):

| Class | Description | Abbreviation |
|-------|-------------|--------------|
| 1 | Cracking | cr |
| 2 | Inclusion | in |
| 3 | Patches | pa |
| 4 | Pitted Surface | ps |
| 5 | Rolled-in Scale | rs |
| 6 | Scratches | sc |

### Folder Structure
```
NEU-CLS/
├── cra/       % 200 images (train: 120, test: 80)
├── in/        % 200 images
├── pa/        % 200 images
├── ps/        % 200 images
├── rs/        % 200 images
└── sc/        % 200 images
```

### Image Format
- Size: 200×200 pixels (will be resized to 224×224)
- Format: JPG, PNG, or BMP
- Color: RGB

### Alternative Datasets
- **NEU-CLS64**: 64×64 version for faster training
- **X-SDD-CLS**: X-steel defect dataset
- **GC10-CLS**: GC10 defect dataset

---

## 5. Quick Start Guide

### Step 1: Add Project to MATLAB Path
```matlab
addpath(genpath('D:/MVRP_Project/TSMMIS_MATLAB'));
```

### Step 2: Verify Data Path
Update `config.dataPath` in `main_TSMMIS.m` to point to your dataset:
```matlab
config.dataPath = fullfile(pwd, 'data', 'NEU-CLS');
```

### Step 3: Run Main Script
```matlab
main_TSMMIS
```

The script will:
1. Load and split data (60% train, 40% test)
2. Initialize ResNet18 encoder
3. Pretrain for 300 epochs
4. Fine-tune with 1-4 labels per class
5. Evaluate and generate t-SNE plot
6. Save results

---

## 6. Detailed Module Description

### 6.1 DataLoader (utils/DataLoader.m)

**Purpose:** Handle data loading, splitting, and multicrop generation

**Key Methods:**
- `loadDataset(rootFolder, trainRatio, fewShotPerClass)` - Load and split data; optional 4th output returns full training pool per class
- `generateMulticrop(images)` - Generate K=5 crops per image
- `applyAugmentation(img)` - Apply random flip, color jitter, blur

**Configuration:**
```matlab
dataLoader = DataLoader([224, 224], 5);  % Image size, K crops
```

### 6.2 ResNet18Encoder (models/ResNet18Encoder.m)

**Purpose:** Build ResNet18 architecture with predictor MLP

**Architecture:**
```
Input (224×224×3)
  → 7×7 conv (64) → BN → ReLU → MaxPool
  → ResBlock1 (64) × 2
  → ResBlock2 (128) × 2
  → ResBlock3 (256) × 2
  → ResBlock4 (512) × 2
  → Global Average Pool
  → FC (512) → ReLU
  → Predictor MLP (512 → 512 → 512)
```

**Usage:**
```matlab
encoder = ResNet18Encoder(6);  % 6 classes
net = encoder.createEncoderWithPredictor();
```

### 6.3 ContrastiveLoss (training/ContrastiveLoss.m)

**Purpose:** Compute L_t&s (teacher-student contrastive loss)

**Mathematical Formulation:**
- Main view (P_s) → Multi-view (Y_k): similarity matrix + cross-entropy
- Multi-view (P_k) → Main view (Y_s): similar computation, Y_s detached
- Temperature τ = 0.1

**Usage:**
```matlab
loss_fn = ContrastiveLoss(0.1);
loss = loss_fn.computeLoss(P_s, Y_k, Y_s, P_k);
```

### 6.4 MMISLoss (training/MMISLoss.m)

**Purpose:** Compute L_MMIS (Multi-Modal Instance Similarity regularization)

**Mathematical Formulation:**
- Pairs: (Y_1,Y_2), (Y_2,Y_3), ..., (Y_K-1, Y_K)
- bmm(Y_i, Y_{i+1}^T) → similarity matrix
- argmax per row → pseudo-labels
- Cross-entropy with negative sign
- Weight: λ × epoch

**Usage:**
```matlab
mmis_fn = MMISLoss(0.001);  % lambda = 0.001
[loss, lossValue] = mmis_fn.computeLoss(Y_k, epoch, K);
```

### 6.5 FineTuner (evaluation/FineTuner.m)

**Purpose:** Few-shot fine-tuning and evaluation

**Two Modes:**
1. **Linear Evaluation:** Backbone frozen, only classifier trained
2. **Full Fine-tuning:** Entire network updated

**Usage:**
```matlab
fineTuner = FineTuner(pretrainedNet, 6);
accuracy = fineTuner.linearEvaluation(fewShotData, testData, 100);
```

---

## 7. Training Configuration

### Hyperparameters Summary

| Parameter | Value | Description |
|-----------|-------|-------------|
| K | 5 | Number of multi-view crops |
| λ | 0.001 | MMIS weight |
| Batch Size | 64 | Training batch size |
| Epochs (Pretrain) | 300 | Self-supervised training |
| Learning Rate | 0.003 | Pretraining LR |
| Momentum | 0.9 | SGD momentum |
| Temperature | 0.1 | Contrastive temperature |
| Fine-tune Epochs | 100 | Few-shot training |
| Fine-tune LR | 0.001 | Fine-tuning LR |
| Num Runs | 100 | Experiments per label count |

### Modifying Configuration

Edit the `config` struct in `main_TSMMIS.m`:

```matlab
% Dataset
config.dataPath = 'path/to/your/data';
config.numClasses = 6;

% Pretraining
config.pretrain.epochs = 300;
config.pretrain.learningRate = 0.003;
config.pretrain.batchSize = 64;
config.pretrain.lambda = 0.001;

% Multicrop
config.K = 5;
config.cropRatioRange = [0.1, 1.0];
```

---

## 8. Running the Pipeline

### Full Pipeline Execution
```matlab
% Run from MATLAB command window
cd D:/MVRP_Project/TSMMIS_MATLAB
main_TSMMIS
```

### Individual Steps

**Step 1: Data Loading**
```matlab
dataLoader = DataLoader([224, 224], 5);
[unlabeled, fewShot, test, trainPool] = dataLoader.loadDataset(dataPath, 0.6, 4);
```

**Step 2: Model Creation**
```matlab
[net, modelInfo] = initializeEncoder(config);  % Uses pretrained resnet18 when available
```

**Step 3: Pretraining**
```matlab
opts = struct('epochs', 300, 'batchSize', 64, 'learningRate', 0.003);
[net, info] = pretrainModel(net, unlabeled, opts);  % Unlabeled feature pretraining
```

**Step 4: Fine-tuning**
```matlab
poolFeatures = applyFeatureTransform(net, extractAllFeatures(net, trainPool));
testFeatures = applyFeatureTransform(net, extractAllFeatures(net, test));
classifier = trainPrototypeClassifier(poolFeatures(sampleIdx, :), poolLabels(sampleIdx), 6);
pred = predictPrototypeClassifier(classifier, testFeatures);
```

**Step 5: Evaluation**
```matlab
acc = mean(pred == testLabels);
fprintf('Accuracy: %.2f%%\n', acc * 100);
```

---

## 9. Understanding the Results

### Output Files
- `models/pretrained_encoder.mat` - Pretrained model weights
- `evaluation/results/evaluation_results.mat` - Accuracy results
- `evaluation/results/tsne_visualization.png` - Feature visualization

### Expected Results (NEU-CLS)
With proper training, expect:
- 1 label/class: ~65-75% accuracy
- 2 labels/class: ~75-85% accuracy  
- 3 labels/class: ~80-90% accuracy
- 4 labels/class: ~85-92% accuracy

### Result Structure
```matlab
results = 
  labels_1: [mean: 0.70, std: 0.05]
  labels_2: [mean: 0.80, std: 0.04]
  labels_3: [mean: 0.85, std: 0.03]
  labels_4: [mean: 0.88, std: 0.03]
```

---

## 10. Troubleshooting

### Common Issues

**1. "Undefined function or variable" errors**
- **Solution:** Ensure all .m files are on MATLAB path
- **Command:** `addpath(genpath('D:/MVRP_Project/TSMMIS_MATLAB'))`

**2. "Out of memory" errors**
- **Solution:** Reduce batch size
- **Code:** `config.pretrain.batchSize = 32;`

**3. "Dataset not found" warning**
- **Solution:** Verify data path in config
- **Check:** `isfolder(config.dataPath)`

**4. Training is very slow**
- **Solution:** Use GPU if available
- **Check:** `gpuDevice`

**5. Low accuracy**
- **Solution:** 
  - Increase training epochs
  - Adjust learning rate
  - Verify data augmentation

### GPU Training
`main_TSMMIS.m` auto-detects CUDA GPU and uses GPU execution for feature extraction when available:
```matlab
if canUseGPU
    gpuInfo = gpuDevice;
    fprintf('Detected GPU: %s\n', gpuInfo.Name);
    config.executionEnvironment = 'gpu';
else
    config.executionEnvironment = 'cpu';
end
```

---

## 11. Ablation Experiments

### 11.1 Effect of K (Number of Crops)
```matlab
K_values = [3, 4, 5, 6];
for K = K_values
    config.K = K;
    % Run training and evaluation
end
```

**Expected:** K=5 gives best results (as per paper)

### 11.2 Effect of MMIS
```matlab
% With MMIS
config.pretrain.lambda = 0.001;

% Without MMIS  
config.pretrain.lambda = 0;
```

### 11.3 Effect of λ
```matlab
lambda_values = [0.0001, 0.001, 0.005, 0.01];
for lambda = lambda_values
    config.pretrain.lambda = lambda;
end
```

### 11.4 Encoder Depth Comparison
```matlab
% ResNet18 (1-1-1-1 blocks)
encoder = ResNet18Encoder(6);

% ResNet34 (2-2-2-2 blocks) - modify ResNet18Encoder.m
```

### 11.5 Blur Robustness Test
```matlab
blurLevels = [0, 1, 2, 3, 4];  % Gaussian sigma
for sigma = blurLevels
    blurredImg = imgaussfilt(testImg, sigma);
    acc = evaluate(net, blurredImg);
end
```

---

## 12. GPU Training Details

### Enabling GPU Acceleration

For proper GPU training in MATLAB, modify the training loop:

```matlab
function [net, info] = pretrainModelGPU(net, dataLoader, data, opts)
    % Use GPU
    net = resetState(net);
    
    for epoch = 1:opts.epochs
        for batch = 1:numBatches
            % Get batch
            [images, ~] = getBatch(data, batch);
            
            % Convert to dlarray and move to GPU
            dlX = dlarray(single(images), 'SSCB');
            
            % Forward pass
            [loss, grad] = dlfeval(@computeLossAndGrad, net, dlX);
            
            % Update
            net = sgdmupdate(net, grad, opts.learningRate);
        end
    end
end

function [loss, grad] = computeLossAndGrad(net, X)
    % Forward pass
    Y = forward(net, X);
    
    % Compute loss
    loss = sum(Y.^2) / numel(Y);
    
    % Compute gradients
    grad = dlgradient(loss, net.Learnables);
end
```

---

## 13. Performance Tips

1. **Use GPU:** GPU training is ~10x faster than CPU
2. **Reduce image size:** Use 64×64 for debugging
3. **Reduce epochs:** Use 50 epochs for testing
4. **Batch processing:** Process multiple images at once
5. **Data caching:** Load entire dataset to memory if RAM allows

---

## 14. Support and Documentation

### Function Help
```matlab
help DataLoader
help ResNet18Encoder
help ContrastiveLoss
help MMISLoss
help FineTuner
```

### References
- Original Paper: TSMMIS (Self-supervised Contrastive Learning with Multi-Modal Instance Similarity)
- NEU-CLS Dataset: https://github.com/cugbr/NEU-CLS

---

## 15. Citation

If you use this code in your research, please cite:

```bibtex
@software{TSMMIS_MATLAB,
  title = {TSMMIS Implementation for MATLAB},
  author = {MVRP Project},
  year = {2026},
  url = {https://github.com/mvrp/tsmmis-matlab}
}
```

---

**End of Documentation**
