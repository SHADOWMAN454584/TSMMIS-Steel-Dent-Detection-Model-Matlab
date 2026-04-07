%% ResNet18 Encoder for TSMMIS
% Implementation of ResNet18 backbone with predictor MLP
% 
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef ResNet18Encoder < handle
    properties
        inputSize = [224, 224, 3];
        numClasses = 6;
        featureDim = 512;
        predictorHiddenDim = 512;
    end
    
    methods
        %% Constructor
        function obj = ResNet18Encoder(numClasses)
            if nargin >= 1
                obj.numClasses = numClasses;
            end
        end
        
        %% Build ResNet18 network using layerGraph
        function lgraph = buildNetwork(obj)
            % Build ResNet18 architecture
            layers = [
                % Initial conv layer
                imageInputLayer(obj.inputSize, 'Name', 'input', 'Normalization', 'none')
                
                % 7x7 conv, stride 2
                convolution2dLayer(7, 64, 'Name', 'conv1', 'Stride', 2, 'Padding', 3, 'BiasLearnRateFactor', 0, 'WeightLearnRateFactor', 0)
                batchNormalizationLayer('Name', 'bn1', 'Epsilon', 1e-5)
                reluLayer('Name', 'relu1')
                
                % 3x3 max pool, stride 2
                maxPooling2dLayer(3, 'Name', 'pool1', 'Stride', 2, 'Padding', 1)
            ];
            
            % Residual block 1 (2 blocks, 64 channels)
            layers = [layers; obj.addResidualBlock(64, 64, 2, 'res2')];
            
            % Residual block 2 (2 blocks, 128 channels, stride 2)
            layers = [layers; obj.addResidualBlock(128, 128, 2, 'res3', 2)];
            
            % Residual block 3 (2 blocks, 256 channels, stride 2)
            layers = [layers; obj.addResidualBlock(256, 256, 2, 'res4', 2)];
            
            % Residual block 4 (2 blocks, 512 channels, stride 2)
            layers = [layers; obj.addResidualBlock(512, 512, 2, 'res5', 2)];
            
            % Global average pooling and fully connected
            layers = [layers;
                globalAveragePooling2dLayer('Name', 'gap')
                fullyConnectedLayer(obj.featureDim, 'Name', 'fc512')
                reluLayer('Name', 'relu_fc')
            ];
            
            % Create layer graph
            lgraph = layerGraph(layers);
        end
        
        %% Add a residual block
        function layers = addResidualBlock(obj, inChannels, outChannels, numBlocks, name, stride)
            if nargin < 6
                stride = 1;
            end
            
            layers = [];
            
            for i = 1:numBlocks
                if i == 1
                    s = stride;
                    inCh = inChannels;
                else
                    s = 1;
                    inCh = outChannels;
                end
                
                blockName = sprintf('%s_%d', name, i);
                
                layers = [layers;
                    convolution2dLayer(3, outChannels, 'Name', [blockName '_conv1'], ...
                        'Stride', s, 'Padding', 1, 'BiasLearnRateFactor', 0, 'WeightLearnRateFactor', 0)
                    batchNormalizationLayer('Name', [blockName '_bn1'], 'Epsilon', 1e-5)
                    reluLayer('Name', [blockName '_relu1'])
                    
                    convolution2dLayer(3, outChannels, 'Name', [blockName '_conv2'], ...
                        'Padding', 1, 'BiasLearnRateFactor', 0, 'WeightLearnRateFactor', 0)
                    batchNormalizationLayer('Name', [blockName '_bn2'], 'Epsilon', 1e-5)
                ];
                
                % Shortcut connection
                if s ~= 1 || inCh ~= outChannels
                    layers = [layers;
                        convolution2dLayer(1, outChannels, 'Name', [blockName '_shortcut_conv'], ...
                            'Stride', s, 'BiasLearnRateFactor', 0, 'WeightLearnRateFactor', 0)
                        batchNormalizationLayer('Name', [blockName '_shortcut_bn'], 'Epsilon', 1e-5)
                    ];
                end
                
                layers = [layers; reluLayer('Name', [blockName '_relu2'])];
            end
        end
        
        %% Create full network with predictor MLP
        function net = createEncoderWithPredictor(obj)
            % Build encoder
            lgraph = obj.buildNetwork();
            
            % Add predictor MLP (2 linear layers)
            predictorLayers = [
                fullyConnectedLayer(obj.predictorHiddenDim, 'Name', 'predictor_fc1')
                reluLayer('Name', 'predictor_relu')
                fullyConnectedLayer(obj.featureDim, 'Name', 'predictor_fc2')
            ];
            
            lgraph = addLayers(lgraph, predictorLayers);
            lgraph = connectLayers(lgraph, 'relu_fc', 'predictor_fc1');
            
            net = dlnetwork(lgraph);
        end
        
        %% Create classification head for fine-tuning
        function net = addClassificationHead(obj, encoderNet, numClasses)
            % Handle both dlnetwork and LayerGraph inputs
            if isa(encoderNet, 'dlnetwork')
                % Extract layer graph from dlnetwork
                lgraph = layerGraph(encoderNet.Layers);
            else
                lgraph = encoderNet;
            end
            
            % Get all layer names
            layerNames = {lgraph.Layers.Name};
            
            % Find the gap layer index
            gapIdx = find(strcmp(layerNames, 'gap'));
            
            if isempty(gapIdx)
                error('Global Average Pooling layer "gap" not found in network');
            end
            
            % Build new network up to and including gap layer
            % Get layers up to gap
            keepLayers = lgraph.Layers(1:gapIdx);
            
            % Add new classification head
            newLayers = [
                fullyConnectedLayer(numClasses, 'Name', 'classifier_fc')
                softmaxLayer('Name', 'softmax')
            ];
            
            % Create new layer graph from scratch
            lgraphNew = layerGraph(keepLayers);
            lgraphNew = addLayers(lgraphNew, newLayers);
            lgraphNew = connectLayers(lgraphNew, 'gap', 'classifier_fc');
            
            net = dlnetwork(lgraphNew);
        end
        
        %% Extract features from encoder (for evaluation)
        function features = extractFeatures(obj, net, images)
            % images: NxHxWxC or NxCxHxW
            
            if size(images, 4) == 3
                % Convert HxWxC to CxHxW
                images = permute(images, [4, 1, 2, 3]);
            end
            
            dlX = dlarray(single(images), 'SSCB');
            
            % Forward pass up to global average pooling
            net = predict(net, dlX);
            
            % Extract features from the feature layer
            features = extractdata(net);
        end
        
        %% Get model parameters for initialization
        function params = getDefaultParams(obj)
            % Xavier/He initialization for weights
            params = struct();
        end
    end
end

%% Alternative: Build ResNet18 using built-in deep network designer
% This function shows how to create ResNet18 using MATLAB's built-in architecture
function net = buildResNet18BuiltIn(numClasses)
    % Load pretrained ResNet18 (if available) and modify
    % Note: Requires Deep Learning Toolbox
    
    net = resnet18;
    
    % Replace the final classification layer
    if nargin >= 1
        net = replaceLayer(net, 'fc1000', fullyConnectedLayer(numClasses));
        net = replaceLayer(net, 'prob', softmaxLayer());
    end
end

%% Quick test function
function testEncoder()
    fprintf('Testing ResNet18 Encoder...\n');
    
    encoder = ResNet18Encoder(6);
    net = encoder.createEncoderWithPredictor();
    
    fprintf('Network created with %d layers\n', numel(net.Layers));
    fprintf('Feature dimension: %d\n', encoder.featureDim);
    
    % Test forward pass
    batchSize = 4;
    X = rand(batchSize, 224, 224, 3, 'single');
    
    dlX = dlarray(X, 'SSCB');
    [Y, state] = forward(net, dlX);
    
    fprintf('Output size: [%s]\n', mat2str(size(Y)));
    fprintf('Encoder test passed!\n');
end