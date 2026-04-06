%% MMIS Regularization Loss for TSMMIS
% Implements the Multi-Modal Instance Similarity (MMIS) regularization
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef MMISLoss < handle
    properties
        lambda = 0.001;  % Weight for MMIS loss
    end
    
    methods
        function obj = MMISLoss(lambda)
            if nargin >= 1
                obj.lambda = lambda;
            end
        end
        
        %% Compute MMIS regularization loss
        % Takes pairs of multi-view encoder outputs and computes consistency loss
        function [loss, lossValue] = computeLoss(obj, Y_k, epoch, K)
            % Inputs:
            %   Y_k - Encoder outputs for multi-view crops (batchSize*K x featureDim)
            %   epoch - Current training epoch (for dynamic weighting)
            %   K - Number of multi-view crops per image
            %
            % Output:
            %   loss - Weighted loss (lambda * epoch * L_MMIS)
            %   lossValue - Raw L_MMIS value
            
            [totalSamples, featureDim] = size(Y_k);
            batchSize = totalSamples / K;
            
            % Reshape to (batchSize, K, featureDim)
            Y_k = reshape(Y_k, [batchSize, K, featureDim]);
            
            % Compute pairwise batch matrix multiplication for adjacent pairs
            % (Y_1, Y_2), (Y_2, Y_3), ..., (Y_K-1, Y_K)
            mmisLoss = 0;
            
            for k = 1:K-1
                Y1 = squeeze(Y_k(:, k, :));      % batchSize x featureDim
                Y2 = squeeze(Y_k(:, k+1, :));    % batchSize x featureDim
                
                % Normalize
                Y1_norm = Y1 ./ (vecnorm(Y1, 2, 2) + 1e-8);
                Y2_norm = Y2 ./ (vecnorm(Y2, 2, 2) + 1e-8);
                
                % Batch matrix multiplication: (batchSize x featureDim) * (featureDim x batchSize)
                % Result: batchSize x batchSize
                M_k = Y1_norm * Y2_norm';
                
                % Extract argmax position per row as pseudo-label
                [~, pseudoLabels] = max(M_k, [], 2);
                pseudoLabels = pseudoLabels - 1;  % Convert to 0-indexed
                
                % Compute cross-entropy of M_k against argmax labels
                ce_k = obj.crossEntropy(M_k, pseudoLabels);
                
                % Negative sign across all K groups
                mmisLoss = mmisLoss - ce_k;
            end
            
            % Average over K-1 pairs
            lossValue = mmisLoss / (K - 1);
            
            % Apply epoch weighting
            loss = obj.lambda * epoch * lossValue;
        end
        
        %% Cross-entropy for MMIS
        function ce = crossEntropy(obj, similarityMatrix, labels)
            batchSize = length(labels);
            
            % Apply softmax
            expSim = exp(similarityMatrix);
            softmaxProb = expSim ./ sum(expSim, 2);
            
            % Get probabilities for correct labels
            correctProbs = zeros(batchSize, 1);
            for i = 1:batchSize
                labelIdx = labels(i) + 1;
                if labelIdx <= size(softmaxProb, 2)
                    correctProbs(i) = softmaxProb(i, labelIdx);
                else
                    correctProbs(i) = 1e-8;
                end
            end
            
            % Cross-entropy
            ce = -mean(log(correctProbs + 1e-8));
        end
    end
end

%% Vectorized version (faster for large batches)
function [loss, lossValue] = computeLossVectorized(obj, Y_k, epoch, K)
    [totalSamples, featureDim] = size(Y_k);
    batchSize = totalSamples / K;
    
    % Reshape to (batchSize, K, featureDim)
    Y_k = reshape(Y_k, [batchSize, K, featureDim]);
    
    % Compute pairwise similarities for all adjacent pairs
    lossValue = 0;
    
    for k = 1:K-1
        Y1 = Y_k(:, k, :);       % batchSize x 1 x featureDim
        Y2 = Y_k(:, k+1, :);     % batchSize x 1 x featureDim
        
        % Squeeze and normalize
        Y1 = squeeze(Y1);
        Y2 = squeeze(Y2);
        
        Y1_norm = Y1 ./ (sqrt(sum(Y1.^2, 2)) + 1e-8);
        Y2_norm = Y2 ./ (sqrt(sum(Y2.^2, 2)) + 1e-8);
        
        % Similarity matrix
        M_k = Y1_norm * Y2_norm';
        
        % Argmax per row
        [~, pseudoLabels] = max(M_k, [], 2);
        pseudoLabels = pseudoLabels - 1;
        
        % Cross-entropy
        expM = exp(M_k);
        probs = expM ./ sum(expM, 2);
        
        % Gather correct probabilities
        idx = sub2ind(size(probs), (1:batchSize)', double(pseudoLabels) + 1);
        correctProbs = probs(idx);
        
        ce_k = -mean(log(correctProbs + 1e-8));
        
        % Negative sign
        lossValue = lossValue - ce_k;
    end
    
    lossValue = lossValue / (K - 1);
    loss = obj.lambda * epoch * lossValue;
end