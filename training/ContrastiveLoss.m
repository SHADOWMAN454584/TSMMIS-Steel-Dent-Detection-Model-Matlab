%% Contrastive Loss (L_t&s) for TSMMIS
% Implements the teacher-student contrastive loss between main view and multi-view crops
%
% Author: TSMMIS-MATLAB Implementation
% Date: April 2026

classdef ContrastiveLoss < handle
    properties
        temperature = 0.1;  % Temperature parameter for softmax
    end
    
    methods
        function obj = ContrastiveLoss(temperature)
            if nargin >= 1
                obj.temperature = temperature;
            end
        end
        
        %% Compute contrastive loss between main view and multi-view crops
        % Main view: P_s (predictor output from encoder output Y_s)
        % Multi-view: Y_k (encoder outputs from multi-view crops)
        % Compute cosine similarity between P_s and Y_k
        function loss = computeLoss(obj, P_s, Y_k, Y_s, P_k)
            % Inputs:
            %   P_s - Predictor output for main view (batchSize x featureDim)
            %   Y_k - Encoder outputs for multi-view crops (batchSize*K x featureDim)
            %   Y_s - Encoder output for main view (batchSize x featureDim)
            %   P_k - Predictor outputs for multi-view crops (batchSize*K x featureDim)
            %
            % Output:
            %   loss - Scalar contrastive loss
            
            [batchSize, featureDim] = size(P_s);
            K = size(Y_k, 1) / batchSize;
            
            % Normalize for cosine similarity
            P_s_norm = P_s ./ (vecnorm(P_s, 2, 2) + 1e-8);
            Y_k_norm = Y_k ./ (vecnorm(Y_k, 2, 2) + 1e-8);
            Y_s_norm = Y_s ./ (vecnorm(Y_s, 2, 2) + 1e-8);
            P_k_norm = P_k ./ (vecnorm(P_k, 2, 2) + 1e-8);
            
            % Compute similarity matrices
            % P_s to Y_k (main -> multi)
            sim_ps_yk = P_s_norm * Y_k_norm';
            sim_ps_yk = sim_ps_yk / obj.temperature;
            
            % P_k to Y_s (multi -> main, with Y_s detached)
            sim_pk_ys = P_k_norm * Y_s_norm';
            sim_pk_ys = sim_pk_ys / obj.temperature;
            
            % Create pseudo-labels: indices [0,1,...,batchSize*K-1] per row
            % For PS -> YK: labels are [0, 1, 2, ..., batchSize-1] repeated for each crop
            labels_ps = repmat(0:batchSize-1, 1, K)';
            
            % For PK -> YS: labels are [0, 1, 2, ..., batchSize-1] repeated K times
            labels_pk = repmat(0:batchSize-1, 1, K)';
            
            % Compute cross-entropy loss for P_s -> Y_k
            loss_ps = obj.crossEntropy(sim_ps_yk, labels_ps);
            
            % Compute cross-entropy loss for P_k -> Y_s (detach Y_s for gradient)
            loss_pk = obj.crossEntropy(sim_pk_ys, labels_pk);
            
            % Average over K crops
            loss = (loss_ps + loss_pk) / (2 * K);
        end
        
        %% Cross-entropy loss with softmax
        function ce = crossEntropy(obj, similarityMatrix, labels)
            % similarityMatrix: (batchSize*K) x (batchSize*K) or (batchSize*K) x batchSize
            % labels: (batchSize*K,) vector of label indices
            
            batchSize = length(labels);
            
            % Apply softmax
            expSim = exp(similarityMatrix);
            softmaxProb = expSim ./ sum(expSim, 2);
            
            % Get probabilities for correct labels
            correctProbs = zeros(batchSize, 1);
            for i = 1:batchSize
                correctProbs(i) = softmaxProb(i, labels(i) + 1);
            end
            
            % Cross-entropy
            ce = -mean(log(correctProbs + 1e-8));
        end
    end
end

%% Helper: batchwise cosine similarity computation
function sim = cosineSimilarityMatrix(X, Y)
    % X: M x D, Y: N x D
    % Returns: M x N similarity matrix
    
    X_norm = X ./ (vecnorm(X, 2, 2) + 1e-8);
    Y_norm = Y ./ (vecnorm(Y, 2, 2) + 1e-8);
    sim = X_norm * Y_norm';
end