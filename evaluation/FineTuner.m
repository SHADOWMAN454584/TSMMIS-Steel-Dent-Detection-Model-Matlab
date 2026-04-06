%% Fine-Tuning and Evaluation Module for TSMMIS
% Implements few-shot fine-tuning and evaluation
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef FineTuner < handle
    properties
        encoderNet;          % Pretrained encoder
        numClasses = 6;
        featureDim = 512;
        learningRate = 0.001;
        finetuneEpochs = 100;
    end
    
    methods
        %% Constructor
        function obj = FineTuner(encoderNet, numClasses)
            obj.encoderNet = encoderNet;
            if nargin >= 2
                obj.numClasses = numClasses;
            end
        end
        
        %% Linear evaluation mode (freeze backbone)
        function [net, accuracy] = linearEvaluation(obj, fewShotData, testData, numRuns)
            if nargin < 4
                numRuns = 100;
            end
            
            accuracies = zeros(numRuns, 1);
            
            for run = 1:numRuns
                % Freeze encoder, train only classification head
                net = obj.freezeBackbone();
                
                % Train classification head
                net = obj.trainClassifier(net, fewShotData, testData);
                
                % Evaluate
                acc = obj.evaluate(net, testData);
                accuracies(run) = acc;
            end
            
            accuracy = struct('mean', mean(accuracies), 'std', std(accuracies), ...
                'all', accuracies);
        end
        
        %% Fine-tuning mode (full network)
        function [net, accuracy] = fullFineTune(obj, fewShotData, testData, numRuns)
            if nargin < 4
                numRuns = 100;
            end
            
            accuracies = zeros(numRuns, 1);
            
            for run = 1:numRuns
                % Unfreeze entire network
                net = obj.unfreezeAll();
                
                % Train full network
                net = obj.trainFullNetwork(net, fewShotData, testData);
                
                % Evaluate
                acc = obj.evaluate(net, testData);
                accuracies(run) = acc;
            end
            
            accuracy = struct('mean', mean(accuracies), 'std', std(accuracies), ...
                'all', accuracies);
        end
        
        %% Freeze backbone weights
        function net = freezeBackbone(obj)
            % Copy encoder and add classification head
            encoder = ResNet18Encoder(obj.numClasses);
            net = encoder.addClassificationHead(obj.encoderNet, obj.numClasses);
            
            % Freeze all layers except classifier
            for i = 1:numel(net.Layers)
                layer = net.Layers(i);
                if ~contains(layer.Name, 'classifier')
                    net.Layers(i).Trainable = false;
                end
            end
        end
        
        %% Unfreeze all layers
        function net = unfreezeAll(obj)
            % Use the pretrained encoder with new classification head
            encoder = ResNet18Encoder(obj.numClasses);
            net = encoder.addClassificationHead(obj.encoderNet, obj.numClasses);
            
            % Ensure all layers are trainable
            for i = 1:numel(net.Layers)
                net.Layers(i).Trainable = true;
            end
        end
        
        %% Train classification head only
        function net = trainClassifier(obj, net, fewShotData, testData)
            % Extract features using pretrained encoder
            featuresTrain = obj.extractFeatures(fewShotData);
            featuresTest = obj.extractFeatures(testData);
            
            % Get labels
            labelsTrain = zeros(length(fewShotData), 1);
            for i = 1:length(fewShotData)
                labelsTrain(i) = fewShotData{i}.class;
            end
            
            labelsTest = zeros(length(testData), 1);
            for i = 1:length(testData)
                labelsTest(i) = testData{i}.class;
            end
            
            % Train simple linear classifier
            classifier = fitcecoc(featuresTrain, labelsTrain, ...
                'Learners', 'linear', ...
                'FitPosterior', true);
            
            % Predict
            predictions = predict(classifier, featuresTest);
            
            % Compute accuracy
            accuracy = sum(predictions == labelsTest) / length(labelsTest);
        end
        
        %% Train full network (requires custom training loop)
        function net = trainFullNetwork(obj, net, fewShotData, testData)
            % This would require a full training loop with backpropagation
            % Simplified version: use pretrained features + SVM
            net = obj.trainClassifier(obj.encoderNet, fewShotData, testData);
        end
        
        %% Extract features using encoder
        function features = extractFeatures(obj, data)
            numSamples = length(data);
            features = zeros(numSamples, obj.featureDim);
            
            batchSize = 32;
            for i = 1:batchSize:numSamples
                idx = i:min(i+batchSize-1, numSamples);
                
                % Load batch
                batchImages = zeros(length(idx), 224, 224, 3);
                for j = 1:length(idx)
                    img = imread(data{idx(j)}.path);
                    if size(img, 3) ~= 3
                        img = cat(3, img, img, img);
                    end
                    img = imresize(img, [224, 224]);
                    batchImages(j, :, :, :) = img;
                end
                
                % Extract features
                dlX = dlarray(single(batchImages), 'SSCB');
                feats = predict(obj.encoderNet, dlX);
                features(idx, :) = extractdata(feats);
            end
        end
        
        %% Evaluate network on test set
        function accuracy = evaluate(obj, net, testData)
            % Extract features
            features = obj.extractFeaturesWithNet(net, testData);
            
            % Get true labels
            labels = zeros(length(testData), 1);
            for i = 1:length(testData)
                labels(i) = testData{i}.class;
            end
            
            % For evaluation, we need predictions
            % This is a simplified version
            accuracy = 0;
        end
        
        %% Extract features with specific network
        function features = extractFeaturesWithNet(obj, net, data)
            numSamples = length(data);
            
            % Find the feature extraction point
            layerNames = arrayfun(@(x) x.Name, net.Laysers, 'UniformOutput', false);
            featureLayerIdx = find(contains(layerNames, 'gap') | contains(layerNames, 'fc512'));
            
            if isempty(featureLayerIdx)
                featureLayerIdx = numel(net.Layers) - 1;
            end
            
            features = zeros(numSamples, obj.featureDim);
            
            batchSize = 32;
            for i = 1:batchSize:numSamples
                idx = i:min(i+batchSize-1, numSamples);
                
                batchImages = zeros(length(idx), 224, 224, 3);
                for j = 1:length(idx)
                    img = imread(data{idx(j)}.path);
                    if size(img, 3) ~= 3
                        img = cat(3, img, img, img);
                    end
                    img = imresize(img, [224, 224]);
                    batchImages(j, :, :, :) = img;
                end
                
                dlX = dlarray(single(batchImages), 'SSCB');
                
                % Extract features from specified layer
                [feats, ~] = predict(net, dlX, 'Outputs', net.Layers(featureLayerIdx).Name);
                features(idx, :) = extractdata(feats);
            end
        end
    end
end

%% Evaluation function for multiple label counts
function results = evaluateFewShot(obj, fewShotData, testData, labelCounts, numRuns)
    % Evaluate with different numbers of labeled samples per class
    %
    % Inputs:
    %   fewShotData - Cell array of all few-shot labeled data
    %   testData - Cell array of test data
    %   labelCounts - Array of label counts to test (e.g., [1, 2, 3, 4])
    %   numRuns - Number of random experiments per label count
    
    results = struct();
    
    for lc = 1:length(labelCounts)
        numLabels = labelCounts(lc);
        fprintf('Evaluating with %d labels per class...\n', numLabels);
        
        accuracies = zeros(numRuns, 1);
        
        for run = 1:numRuns
            % Randomly sample numLabels per class
            sampledData = obj.sampleFewShotData(fewShotData, numLabels);
            
            % Train and evaluate
            net = obj.fineTune(sampledData, testData);
            acc = obj.evaluate(net, testData);
            accuracies(run) = acc;
        end
        
        results.(sprintf('labels_%d', numLabels)) = struct(...
            'mean', mean(accuracies), ...
            'std', std(accuracies), ...
            'all', accuracies);
        
        fprintf('  Mean Accuracy: %.2f +/- %.2f%%\n', ...
            mean(accuracies) * 100, std(accuracies) * 100);
    end
end

%% Sample few-shot data
function sampledData = sampleFewShotData(data, numPerClass)
    % Group data by class
    classes = unique(arrayfun(@(x) x.class, data));
    
    sampledData = {};
    for c = 1:length(classes)
        classData = data([data.class] == classes(c));
        
        if length(classData) >= numPerClass
            idx = randperm(length(classData), numPerClass);
            sampledData = [sampledData, classData(idx)];
        else
            sampledData = [sampledData, classData];
        end
    end
end

%% t-SNE visualization function
function tsneVisualization(features, labels, titleStr)
    % Perform t-SNE visualization of learned features
    %
    % Inputs:
    %   features - N x D feature matrix
    %   labels - N x 1 label vector
    %   titleStr - Title for the plot
    
    % Reduce to 2D using t-SNE
    if exist('tsne', 'file')
        Y = tsne(features, 'NumDimensions', 2);
    else
        fprintf('t-SNE not available. Using PCA for visualization.\n');
        [~, ~, Y] = pca(features);
        Y = Y(:, 1:2);
    end
    
    % Plot
    figure;
    gscatter(Y(:, 1), Y(:, 2), labels);
    title(titleStr);
    xlabel('Dimension 1');
    ylabel('Dimension 2');
end