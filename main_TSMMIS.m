%% TSMMIS Main Execution Script
% Complete pipeline for self-supervised few-shot learning
%
% This script runs the TSMMIS pipeline:
%   1. Data loading and preprocessing
%   2. Encoder initialization
%   3. Self-supervised feature pretraining on unlabeled data
%   4. Few-shot fine-tuning/evaluation
%   5. t-SNE visualization
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

%% Clear workspace
clear; close all; clc;

%% Set random seed for reproducibility
rng(42);

%% ===================== CONFIGURATION =====================
% Dataset configuration
config.dataPath = fullfile(pwd, 'data', 'NEU-CLS');
config.numClasses = 6;
config.trainRatio = 0.6;
config.imageSize = [224, 224];

% Multicrop configuration
config.K = 5;
config.cropRatioRange = [0.1, 1.0];

% Pretraining configuration
config.pretrain.epochs = 300;
config.pretrain.batchSize = 64;
config.pretrain.learningRate = 0.003;
config.pretrain.momentum = 0.9;
config.pretrain.weightDecay = 1e-4;
config.pretrain.lambda = 0.001;
config.pretrain.temperature = 0.1;
config.pretrain.pcaDim = 256;

% Fine-tuning configuration
config.finetune.numRuns = 100;
config.finetune.labelCounts = [1, 2, 3, 4];
config.finetune.selfTrainIterations = 2;
config.finetune.selfTrainMargin = 0.08;

% Paths
config.savePath = fullfile(pwd, 'models');
config.resultsPath = fullfile(pwd, 'evaluation', 'results');

if ~isfolder(config.savePath)
    mkdir(config.savePath);
end
if ~isfolder(config.resultsPath)
    mkdir(config.resultsPath);
end

% Execution environment
useGPU = false;
if exist('canUseGPU', 'file') == 2
    useGPU = canUseGPU;
end
config.executionEnvironment = 'cpu';
if useGPU
    config.executionEnvironment = 'gpu';
    gpuInfo = gpuDevice;
    fprintf('Detected GPU: %s\n', gpuInfo.Name);
else
    fprintf('GPU not available. Using CPU execution.\n');
end

%% ===================== STEP 1: DATA PREPARATION =====================
fprintf('=== STEP 1: Data Preparation ===\n');

% Initialize data loader
dataLoader = DataLoader(config.imageSize, config.K);

% Load dataset
if isfolder(config.dataPath)
    [unlabeledData, fewShotData, testData, trainPoolData] = dataLoader.loadDataset( ...
        config.dataPath, config.trainRatio, max(config.finetune.labelCounts));
    fewShotPool = trainPoolData;
else
    fprintf('Dataset not found at %s\n', config.dataPath);
    fprintf('Creating synthetic data for testing...\n');
    [unlabeledData, fewShotData, testData, trainPoolData] = createSyntheticData(config);
    fewShotPool = trainPoolData;
end

fprintf('Data preparation complete!\n');
fprintf('  Unlabeled: %d\n  Few-shot (reserved): %d\n  Few-shot pool: %d\n  Test: %d\n', ...
    length(unlabeledData), length(fewShotData), length(fewShotPool), length(testData));

%% ===================== STEP 2: MODEL INITIALIZATION =====================
fprintf('\n=== STEP 2: Model Initialization ===\n');

[net, modelInfo] = initializeEncoder(config);

fprintf('Network created: %s\n', modelInfo.name);
fprintf('Feature dimension: %d\n', modelInfo.featureDim);
fprintf('Execution environment: %s\n', modelInfo.executionEnvironment);

%% ===================== STEP 3: SELF-SUPERVISED PRETRAINING =====================
fprintf('\n=== STEP 3: Self-Supervised Pretraining ===\n');

trainOpts = struct();
trainOpts.learningRate = config.pretrain.learningRate;
trainOpts.momentum = config.pretrain.momentum;
trainOpts.weightDecay = config.pretrain.weightDecay;
trainOpts.epochs = config.pretrain.epochs;
trainOpts.batchSize = config.pretrain.batchSize;
trainOpts.K = config.K;
trainOpts.lambda = config.pretrain.lambda;
trainOpts.temperature = config.pretrain.temperature;
trainOpts.numClasses = config.numClasses;
trainOpts.pcaDim = config.pretrain.pcaDim;

fprintf('Starting pretraining for %d epochs...\n', config.pretrain.epochs);
fprintf('Batch size: %d, Learning rate: %.4f\n', config.pretrain.batchSize, config.pretrain.learningRate);

[net, trainInfo] = pretrainModel(net, unlabeledData, trainOpts);

modelPath = fullfile(config.savePath, 'pretrained_encoder.mat');
save(modelPath, 'net', 'trainInfo', 'modelInfo', '-v7.3');
fprintf('Model saved to %s\n', modelPath);

%% ===================== STEP 4: FEW-SHOT FINE-TUNING =====================
fprintf('\n=== STEP 4: Few-Shot Fine-Tuning ===\n');

if isfile(modelPath)
    loaded = load(modelPath, 'net');
    net = loaded.net;
    fprintf('Loaded pretrained model from %s\n', modelPath);
else
    error('Pretrained model not found at %s. Run Step 3 first.', modelPath);
end

fprintf('Precomputing feature bank for few-shot evaluation...\n');
poolFeaturesRaw = extractAllFeatures(net, fewShotPool);
unlabeledFeaturesRaw = extractAllFeatures(net, unlabeledData);
testFeaturesRaw = extractAllFeatures(net, testData);

poolFeatures = applyFeatureTransform(net, poolFeaturesRaw);
unlabeledFeatures = applyFeatureTransform(net, unlabeledFeaturesRaw);
testFeatures = applyFeatureTransform(net, testFeaturesRaw);

poolLabels = getDataLabels(fewShotPool);
testLabels = getDataLabels(testData);

results = struct();
for lc = 1:length(config.finetune.labelCounts)
    numLabels = config.finetune.labelCounts(lc);
    fprintf('\nEvaluating with %d label(s) per class...\n', numLabels);
    
    accuracies = zeros(config.finetune.numRuns, 1);
    
    for run = 1:config.finetune.numRuns
        sampledIdx = sampleFewShotIndices(poolLabels, numLabels, config.numClasses);
        supportFeatures = poolFeatures(sampledIdx, :);
        supportLabels = poolLabels(sampledIdx);
        
        classifier = trainPrototypeClassifier(supportFeatures, supportLabels, config.numClasses);
        classifier = adaptPrototypeClassifier( ...
            classifier, unlabeledFeatures, ...
            config.finetune.selfTrainIterations, ...
            config.finetune.selfTrainMargin);
        
        predictions = predictPrototypeClassifier(classifier, testFeatures);
        accuracies(run) = mean(predictions == testLabels);
    end
    
    fieldName = sprintf('labels_%d', numLabels);
    results.(fieldName) = struct( ...
        'mean', mean(accuracies), ...
        'std', std(accuracies), ...
        'all', accuracies);
    
    fprintf('  Accuracy: %.2f +/- %.2f%%\n', mean(accuracies) * 100, std(accuracies) * 100);
end

%% ===================== STEP 5: EVALUATION RESULTS =====================
fprintf('\n=== STEP 5: Evaluation Results ===\n');
fprintf('\n%-20s | %-15s | %-15s\n', 'Configuration', 'Mean Accuracy', 'Std Dev');
fprintf('%s\n', repmat('-', 1, 55));
for lc = 1:length(config.finetune.labelCounts)
    fieldName = sprintf('labels_%d', config.finetune.labelCounts(lc));
    fprintf('%-20s | %-14.2f%% | %-14.2f%%\n', ...
        fieldName, results.(fieldName).mean * 100, results.(fieldName).std * 100);
end

%% ===================== STEP 6: FEATURE VISUALIZATION =====================
fprintf('\n=== STEP 6: Feature Visualization ===\n');
fprintf('Extracting features for t-SNE visualization...\n');

tsneLabels = testLabels;
tsneInput = testFeatures;
if exist('tsne', 'file') == 2 && size(tsneInput, 1) > 2
    tsnePlot = tsne(tsneInput, 'NumDimensions', 2, 'Perplexity', 30, 'Standardize', true);
else
    [~, score] = pca(tsneInput);
    tsnePlot = score(:, 1:2);
end

tsneDataPath = fullfile(config.resultsPath, 'tsne_embedding.mat');
save(tsneDataPath, 'tsnePlot', 'tsneLabels');

figPath = fullfile(config.resultsPath, 'tsne_visualization.png');
[didSaveFigure, saveMessage] = saveTsneFigure(tsnePlot, tsneLabels, figPath);
if didSaveFigure
    fprintf('t-SNE plot saved to %s\n', figPath);
else
    warning('t-SNE image export skipped: %s', saveMessage);
end

%% ===================== SAVE RESULTS =====================
resultsPath = fullfile(config.resultsPath, 'evaluation_results.mat');
save(resultsPath, 'results', 'config', 'modelInfo', 'trainInfo');
fprintf('\nResults saved to %s\n', resultsPath);

fprintf('\n=== TSMMIS Pipeline Complete! ===\n');


%% ===================== HELPER FUNCTIONS =====================
function [unlabeledData, fewShotData, testData, trainPoolData] = createSyntheticData(config)
    numImages = 300;  % Per class
    unlabeledData = {};
    fewShotData = {};
    testData = {};
    trainPoolData = {};
    
    fewShotPerClass = max(config.finetune.labelCounts);
    trainCount = floor(numImages * config.trainRatio);
    
    for c = 1:config.numClasses
        classDir = fullfile(config.dataPath, sprintf('class_%d', c));
        if ~isfolder(classDir)
            mkdir(classDir);
        end
        
        for i = 1:numImages
            img = uint8(rand(config.imageSize(1), config.imageSize(2), 3) * 255);
            imgPath = fullfile(classDir, sprintf('img_%04d.png', i));
            if ~isfile(imgPath)
                imwrite(img, imgPath);
            end
            
            sample = struct('path', imgPath, 'class', c - 1);
            
            if i <= trainCount
                trainPoolData{end+1} = sample;
                if i <= fewShotPerClass
                    fewShotData{end+1} = sample;
                else
                    unlabeledData{end+1} = sample;
                end
            else
                testData{end+1} = sample;
            end
        end
    end
end

function [encoderModel, info] = initializeEncoder(config)
    if exist('resnet18', 'file') == 2
        backbone = resnet18;
        inputSize = backbone.Layers(1).InputSize(1:2);
        encoderModel = struct( ...
            'backend', 'resnet18', ...
            'network', backbone, ...
            'featureLayer', 'pool5', ...
            'featureDim', 512, ...
            'inputSize', inputSize, ...
            'executionEnvironment', config.executionEnvironment, ...
            'transform', struct());
        
        info = struct();
        info.name = 'ImageNet-pretrained ResNet18';
        info.featureDim = encoderModel.featureDim;
        info.executionEnvironment = config.executionEnvironment;
    else
        warning(['resnet18 is unavailable. Install Deep Learning Toolbox Model for ResNet-18 Network ' ...
                 'for best few-shot performance. Falling back to project encoder.']);
        
        encoder = ResNet18Encoder(config.numClasses);
        fallbackNet = encoder.createEncoderWithPredictor();
        
        encoderModel = struct( ...
            'backend', 'dlnetwork', ...
            'network', fallbackNet, ...
            'featureLayer', '', ...
            'featureDim', config.pretrain.pcaDim, ...
            'inputSize', config.imageSize, ...
            'executionEnvironment', config.executionEnvironment, ...
            'transform', struct());
        
        info = struct();
        info.name = 'Fallback project encoder';
        info.featureDim = encoderModel.featureDim;
        info.executionEnvironment = config.executionEnvironment;
    end
end

function [encoderModel, info] = pretrainModel(encoderModel, unlabeledData, opts)
    if isempty(unlabeledData)
        error('No unlabeled samples available for pretraining.');
    end
    
    features = extractAllFeatures(encoderModel, unlabeledData);
    features = normalizeRows(features);
    
    pcaDim = min(opts.pcaDim, size(features, 2));
    [coeff, score, ~, ~, explained, mu] = pca(features, 'NumComponents', pcaDim);
    score = normalizeRows(score);
    
    [~, centers] = kmeans(score, opts.numClasses, ...
        'Replicates', 5, 'MaxIter', 100, 'Display', 'off');
    centers = normalizeRows(centers);
    
    info = struct();
    info.epoch = (1:opts.epochs)';
    info.loss = zeros(opts.epochs, 1);
    
    for epoch = 1:opts.epochs
        similarity = score * centers';
        [maxSimilarity, assignments] = max(similarity, [], 2);
        
        for c = 1:opts.numClasses
            classMembers = score(assignments == c, :);
            if ~isempty(classMembers)
                centers(c, :) = normalizeRows(mean(classMembers, 1));
            end
        end
        
        info.loss(epoch) = mean(1 - maxSimilarity);
        
        if mod(epoch, 10) == 0
            fprintf('Epoch [%d/%d], Loss: %.4f\n', epoch, opts.epochs, info.loss(epoch));
        end
    end
    
    encoderModel.transform.mu = mu;
    encoderModel.transform.coeff = coeff;
    encoderModel.transform.prototypeCenters = centers;
    encoderModel.transform.explainedVariance = sum(explained);
end

function features = extractAllFeatures(encoderModel, data)
    numSamples = length(data);
    if numSamples == 0
        features = zeros(0, encoderModel.featureDim, 'single');
        return;
    end
    
    batchSize = 32;
    targetSize = encoderModel.inputSize;
    features = zeros(numSamples, encoderModel.featureDim, 'single');
    
    for i = 1:batchSize:numSamples
        idx = i:min(i + batchSize - 1, numSamples);
        batchImages = zeros(targetSize(1), targetSize(2), 3, length(idx), 'single');
        
        for j = 1:length(idx)
            img = imread(data{idx(j)}.path);
            if size(img, 3) ~= 3
                img = cat(3, img, img, img);
            end
            img = imresize(img, targetSize);
            batchImages(:, :, :, j) = single(img);
        end
        
        if strcmp(encoderModel.backend, 'resnet18')
            batchFeatures = activations(encoderModel.network, batchImages, encoderModel.featureLayer, ...
                'OutputAs', 'rows', 'ExecutionEnvironment', encoderModel.executionEnvironment);
            features(idx, :) = single(batchFeatures);
        else
            dlX = dlarray(batchImages, 'SSCB');
            if strcmp(encoderModel.executionEnvironment, 'gpu') && exist('canUseGPU', 'file') == 2 && canUseGPU
                dlX = gpuArray(dlX);
            end
            
            rawFeatures = predict(encoderModel.network, dlX);
            rawFeatures = gather(extractdata(rawFeatures));
            rawFeatures = squeeze(rawFeatures);
            if ndims(rawFeatures) > 2
                rawFeatures = reshape(rawFeatures, [], length(idx))';
            end
            
            if size(rawFeatures, 1) ~= length(idx)
                rawFeatures = rawFeatures';
            end
            
            if size(rawFeatures, 2) > encoderModel.featureDim
                rawFeatures = rawFeatures(:, 1:encoderModel.featureDim);
            elseif size(rawFeatures, 2) < encoderModel.featureDim
                rawFeatures = padarray(rawFeatures, [0, encoderModel.featureDim - size(rawFeatures, 2)], 0, 'post');
            end
            
            features(idx, :) = single(rawFeatures);
        end
    end
end

function transformed = applyFeatureTransform(encoderModel, features)
    if isempty(features)
        transformed = features;
        return;
    end
    
    transformed = normalizeRows(features);
    
    if isfield(encoderModel, 'transform') && isfield(encoderModel.transform, 'coeff') ...
            && ~isempty(encoderModel.transform.coeff)
        centered = bsxfun(@minus, double(transformed), encoderModel.transform.mu);
        projected = centered * encoderModel.transform.coeff;
        transformed = normalizeRows(projected);
    end
end

function labels = getDataLabels(data)
    labels = zeros(length(data), 1);
    for i = 1:length(data)
        labels(i) = data{i}.class + 1;
    end
end

function sampledIdx = sampleFewShotIndices(poolLabels, numPerClass, numClasses)
    sampledIdx = zeros(numPerClass * numClasses, 1);
    cursor = 1;
    
    for c = 1:numClasses
        classIdx = find(poolLabels == c);
        if length(classIdx) < numPerClass
            error('Class %d has only %d samples; cannot sample %d.', c, length(classIdx), numPerClass);
        end
        
        selected = classIdx(randperm(length(classIdx), numPerClass));
        sampledIdx(cursor:cursor + numPerClass - 1) = selected(:);
        cursor = cursor + numPerClass;
    end
end

function classifier = trainPrototypeClassifier(features, labels, numClasses)
    features = normalizeRows(features);
    featureDim = size(features, 2);
    prototypes = zeros(numClasses, featureDim);
    
    for c = 1:numClasses
        classMask = labels == c;
        if ~any(classMask)
            error('Support set is missing class %d.', c);
        end
        prototypes(c, :) = mean(features(classMask, :), 1);
    end
    
    classifier = struct();
    classifier.numClasses = numClasses;
    classifier.prototypes = normalizeRows(prototypes);
    classifier.supportFeatures = features;
    classifier.supportLabels = labels;
end

function classifier = adaptPrototypeClassifier(classifier, unlabeledFeatures, numIterations, marginThreshold)
    if isempty(unlabeledFeatures) || numIterations <= 0
        return;
    end
    
    unlabeledFeatures = normalizeRows(unlabeledFeatures);
    
    for iter = 1:numIterations
        similarity = unlabeledFeatures * classifier.prototypes';
        [sortedScores, sortedIdx] = sort(similarity, 2, 'descend');
        predictedClass = sortedIdx(:, 1);
        confidenceMargin = sortedScores(:, 1) - sortedScores(:, 2);
        
        confidentMask = confidenceMargin >= marginThreshold;
        if ~any(confidentMask)
            break;
        end
        
        updatedAnyClass = false;
        for c = 1:classifier.numClasses
            classMask = confidentMask & predictedClass == c;
            if any(classMask)
                classMean = mean(unlabeledFeatures(classMask, :), 1);
                classifier.prototypes(c, :) = normalizeRows( ...
                    0.7 * classifier.prototypes(c, :) + 0.3 * classMean);
                updatedAnyClass = true;
            end
        end
        
        if ~updatedAnyClass
            break;
        end
    end
end

function predictions = predictPrototypeClassifier(classifier, features)
    features = normalizeRows(features);
    protoSimilarity = features * classifier.prototypes';
    
    supportSimilarity = features * classifier.supportFeatures';
    supportClassSimilarity = -inf(size(features, 1), classifier.numClasses);
    for c = 1:classifier.numClasses
        classMask = classifier.supportLabels == c;
        if any(classMask)
            supportClassSimilarity(:, c) = max(supportSimilarity(:, classMask), [], 2);
        end
    end
    
    combinedSimilarity = 0.75 * protoSimilarity + 0.25 * supportClassSimilarity;
    [~, predictions] = max(combinedSimilarity, [], 2);
end

function X = normalizeRows(X)
    if isempty(X)
        return;
    end
    
    rowNorms = sqrt(sum(X.^2, 2));
    rowNorms(rowNorms < 1e-12) = 1;
    X = X ./ rowNorms;
end

function [didSave, message] = saveTsneFigure(tsnePlot, labels, figPath)
    didSave = false;
    message = '';
    
    if ~feature('ShowFigureWindows')
        message = 'Figure windows are disabled in this MATLAB session.';
        return;
    end
    
    fig = [];
    try
        fig = figure('Visible', 'off');
        ax = axes(fig);
        gscatter(ax, tsnePlot(:, 1), tsnePlot(:, 2), labels);
        title(ax, 't-SNE Visualization of Learned Features');
        xlabel(ax, 'Dimension 1');
        ylabel(ax, 'Dimension 2');
        drawnow;
        
        if exist('exportgraphics', 'file') == 2
            exportgraphics(ax, figPath, 'Resolution', 300);
        else
            saveas(fig, figPath);
        end
        
        didSave = true;
    catch ME
        message = ME.message;
    end
    
    if ~isempty(fig) && isgraphics(fig)
        close(fig);
    end
end
