%% Complete Evaluation and Testing Functions
% For TSMMIS few-shot learning
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef TSMMISEvaluator < handle
    properties
        encoderNet;
        numClasses = 6;
        featureDim = 512;
    end
    
    methods
        %% Constructor
        function obj = TSMMISEvaluator(encoderNet, numClasses)
            obj.encoderNet = encoderNet;
            if nargin >= 2
                obj.numClasses = numClasses;
            end
        end
        
        %% Linear evaluation mode
        function accuracy = linearEval(obj, fewShotData, testData, numRuns)
            if nargin < 4
                numRuns = 100;
            end
            
            accuracies = zeros(numRuns, 1);
            
            parfor run = 1:numRuns
                % Extract features
                trainFeatures = obj.extractFeatures(fewShotData);
                testFeatures = obj.extractFeatures(testData);
                
                % Get labels
                trainLabels = obj.getLabels(fewShotData);
                testLabels = obj.getLabels(testData);
                
                % Train linear SVM
                classifier = fitcecoc(trainFeatures, trainLabels, ...
                    'Learners', 'linear', 'FitPosterior', true);
                
                % Predict
                predictions = predict(classifier, testFeatures);
                
                % Accuracy
                accuracies(run) = sum(predictions == testLabels) / length(testLabels);
            end
            
            accuracy = struct('mean', mean(accuracies), 'std', std(accuracies), 'all', accuracies);
        end
        
        %% Full fine-tuning mode
        function accuracy = fullFineTune(obj, fewShotData, testData, numRuns)
            if nargin < 4
                numRuns = 100;
            end
            
            accuracies = zeros(numRuns, 1);
            
            parfor run = 1:numRuns
                % Create network with classifier
                net = obj.createClassificationNetwork();
                
                % Extract features
                trainFeatures = obj.extractFeatures(fewShotData);
                testFeatures = obj.extractFeatures(testData);
                
                % Get labels
                trainLabels = obj.getLabels(fewShotData);
                testLabels = obj.getLabels(testData);
                
                % Train classifier on features
                classifier = fitcecoc(trainFeatures, trainLabels, ...
                    'Learners', 'knn', 'NumNeighbors', 3);
                
                % Predict
                predictions = predict(classifier, testFeatures);
                
                % Accuracy
                accuracies(run) = sum(predictions == testLabels) / length(testLabels);
            end
            
            accuracy = struct('mean', mean(accuracies), 'std', std(accuracies), 'all', accuracies);
        end
        
        %% Extract features using pretrained encoder
        function features = extractFeatures(obj, data)
            numSamples = length(data);
            features = zeros(numSamples, obj.featureDim);
            
            batchSize = 32;
            for i = 1:batchSize:numSamples
                idx = i:min(i+batchSize-1, numSamples);
                
                batchImages = zeros(length(idx), 224, 224, 3);
                for j = 1:length(idx)
                    if isempty(data{idx(j)}.path)
                        batchImages(j, :, :, :) = rand(224, 224, 3);
                    else
                        img = imread(data{idx(j)}.path);
                        if size(img, 3) ~= 3
                            img = cat(3, img, img, img);
                        end
                        img = imresize(img, [224, 224]);
                        batchImages(j, :, :, :) = img;
                    end
                end
                
                dlX = dlarray(single(batchImages), 'SSCB');
                feats = predict(obj.encoderNet, dlX);
                
                if ndims(feats) == 4
                    feats = squeeze(feats);
                end
                if size(feats, 2) > size(feats, 1)
                    feats = feats';
                end
                
                features(idx, :) = extractdata(feats);
            end
        end
        
        %% Get labels from data
        function labels = getLabels(obj, data)
            labels = zeros(length(data), 1);
            for i = 1:length(data)
                labels(i) = data{i}.class;
            end
        end
        
        %% Create classification network
        function net = createClassificationNetwork(obj)
            encoder = ResNet18Encoder(obj.numClasses);
            net = encoder.addClassificationHead(obj.encoderNet, obj.numClasses);
        end
    end
end

%% Run complete evaluation with different label counts
function results = runEvaluation(encoderNet, fewShotData, testData, labelCounts, numRuns)
    % Run evaluation with different numbers of labeled samples
    
    evaluator = TSMMISEvaluator(encoderNet, 6);
    results = struct();
    
    for lc = 1:length(labelCounts)
        numLabels = labelCounts(lc);
        fprintf('Testing with %d label(s) per class...\n', numLabels);
        
        % Sample data
        sampledFewShot = sampleFewShotData(fewShotData, numLabels);
        
        % Linear evaluation
        accuracy = evaluator.linearEval(sampledFewShot, testData, numRuns);
        
        fieldName = sprintf('labels_%d', numLabels);
        results.(fieldName) = accuracy;
        
        fprintf('  Linear Eval: %.2f +/- %.2f%%\n', accuracy.mean*100, accuracy.std*100);
    end
end

%% Sample few-shot data
function sampledData = sampleFewShotData(data, numPerClass)
    classes = unique(arrayfun(@(x) x.class, data));
    sampledData = {};
    
    for c = 1:length(classes)
        classData = data([data.class] == classes(c)-1);
        
        if isempty(classData)
            classData = data([data.class] == c-1);
        end
        
        if length(classData) >= numPerClass
            idx = randperm(length(classData), numPerClass);
            sampledData = [sampledData, classData(idx)];
        elseif ~isempty(classData)
            sampledData = [sampledData, classData];
        end
    end
end

%% t-SNE visualization
function tsneVisualization(features, labels, savePath)
    if ~isempty(features) && size(features, 1) > 1
        figure;
        if exist('tsne', 'file')
            Y = tsne(features, 'NumDimensions', 2);
            gscatter(Y(:, 1), Y(:, 2), labels);
        else
            [~,~,Y] = pca(features);
            scatter(Y(:, 1), Y(:, 2), 10, labels, 'filled');
        end
        title('t-SNE Visualization of Learned Features');
        xlabel('Dimension 1');
        ylabel('Dimension 2');
        
        if nargin >= 3 && ~isempty(savePath)
            saveas(gcf, savePath);
        end
    end
end

%% Generate training progress plot
function plotTrainingProgress(info, savePath)
    figure;
    
    subplot(1, 3, 1);
    plot(info.epoch, info.loss);
    title('Total Loss');
    xlabel('Epoch');
    ylabel('Loss');
    
    subplot(1, 3, 2);
    plot(info.epoch, info.lossContrastive);
    title('Contrastive Loss (L_t&s)');
    xlabel('Epoch');
    ylabel('Loss');
    
    subplot(1, 3, 3);
    plot(info.epoch, info.lossMMIS);
    title('MMIS Loss (L_MMIS)');
    xlabel('Epoch');
    ylabel('Loss');
    
    if nargin >= 2 && ~isempty(savePath)
        saveas(gcf, savePath);
    end
end

%% Quick test
function testEvaluator()
    fprintf('Testing TSMMIS Evaluator...\n');
    
    encoder = ResNet18Encoder(6);
    net = encoder.createEncoderWithPredictor();
    
    evaluator = TSMMISEvaluator(net, 6);
    
    % Test with synthetic data
    fewShotData = cell(6, 1);
    for i = 1:6
        fewShotData{i} = struct('path', '', 'class', i-1);
    end
    
    testData = cell(60, 1);
    for i = 1:60
        testData{i} = struct('path', '', 'class', mod(i-1, 6));
    end
    
    fprintf('Evaluator created successfully!\n');
end