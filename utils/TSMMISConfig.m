%% TSMMIS Configuration File
% All hyperparameters and paths in one place
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef TSMMISConfig
    properties
        % Dataset
        dataPath = 'data/NEU-CLS';
        numClasses = 6;
        trainRatio = 0.6;
        imageSize = [224, 224];
        
        % Multicrop
        K = 5;
        cropRatioRange = [0.1, 1.0];
        mainCropRatio = 0.8;
        
        % Pretraining
        pretrain = struct();
        
        % Fine-tuning
        finetune = struct();
        
        % Model
        encoder = 'ResNet18';
        featureDim = 512;
        predictorHiddenDim = 512;
        
        % Paths
        savePath = 'models';
        resultsPath = 'evaluation/results';
    end
    
    methods
        %% Constructor with defaults
        function obj = TSMMISConfig()
            % Pretraining defaults
            obj.pretrain.epochs = 300;
            obj.pretrain.batchSize = 64;
            obj.pretrain.learningRate = 0.003;
            obj.pretrain.momentum = 0.9;
            obj.pretrain.weightDecay = 1e-4;
            obj.pretrain.lambda = 0.001;
            obj.pretrain.temperature = 0.1;
            
            % Fine-tuning defaults
            obj.finetune.epochs = 100;
            obj.finetune.learningRate = 0.001;
            obj.finetune.numRuns = 100;
            obj.finetune.labelCounts = [1, 2, 3, 4];
        end
        
        %% Load from file
        function obj = loadFromFile(obj, configFile)
            if exist(configFile, 'file')
                loaded = load(configFile);
                fields = fieldnames(loaded);
                for i = 1:length(fields)
                    obj.(fields{i}) = loaded.(fields{i});
                end
            end
        end
        
        %% Save to file
        function saveToFile(obj, configFile)
            save(configFile, 'obj');
        end
        
        %% Display configuration
        function displayConfig(obj)
            fprintf('\n=== TSMMIS Configuration ===\n');
            fprintf('Dataset: %s\n', obj.dataPath);
            fprintf('Classes: %d\n', obj.numClasses);
            fprintf('Image Size: [%d, %d]\n', obj.imageSize(1), obj.imageSize(2));
            fprintf('\n--- Pretraining ---\n');
            fprintf('Epochs: %d\n', obj.pretrain.epochs);
            fprintf('Batch Size: %d\n', obj.pretrain.batchSize);
            fprintf('Learning Rate: %.4f\n', obj.pretrain.learningRate);
            fprintf('Momentum: %.2f\n', obj.pretrain.momentum);
            fprintf('Lambda (MMIS): %.4f\n', obj.pretrain.lambda);
            fprintf('Temperature: %.2f\n', obj.pretrain.temperature);
            fprintf('K (crops): %d\n', obj.K);
            fprintf('\n--- Fine-tuning ---\n');
            fprintf('Epochs: %d\n', obj.finetune.epochs);
            fprintf('Learning Rate: %.4f\n', obj.finetune.learningRate);
            fprintf('Num Runs: %d\n', obj.finetune.numRuns);
            fprintf('Label Counts: [%s]\n', num2str(obj.finetune.labelCounts));
            fprintf('========================\n\n');
        end
    end
end

%% Quick configuration test
function testConfig()
    config = TSMMISConfig();
    config.displayConfig();
    fprintf('Configuration test passed!\n');
end