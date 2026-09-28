  function numericValue = sanitizeKpiValue(rawValue)
% sanitizeKpiValue
% Ensures KPI values are numeric.
% Replaces NaN / N/A / NA / missing / empty / text with NaN.
%
% Output is ALWAYS a scalar double (NaN if invalid).

    %------------------------------------------------------
    % Case 0: empty input ([])
    %------------------------------------------------------
    if isempty(rawValue)
        numericValue = NaN;
        return;
    end

    %------------------------------------------------------
    % Case 1: missing (string, datetime, categorical, etc.)
    %------------------------------------------------------
    if ismissing(rawValue)
        numericValue = NaN;
        return;
    end

    %------------------------------------------------------
    % Case 2: numeric
    %------------------------------------------------------
    if isnumeric(rawValue)
        if isscalar(rawValue) && ~isnan(rawValue)
            numericValue = double(rawValue);
        else
            numericValue = NaN;
        end
        return;
    end

    %------------------------------------------------------
    % Case 3: logical
    %------------------------------------------------------
    if islogical(rawValue)
        numericValue = double(rawValue); % true -> 1, false -> 0
        return;
    end

    %------------------------------------------------------
    % Case 4: datetime (KPI date makes no sense -> NaN)
    %------------------------------------------------------
    if isa(rawValue, 'datetime')
        numericValue = NaN;
        return;
    end

    %------------------------------------------------------
    % Case 5: string or char
    %------------------------------------------------------
    if isstring(rawValue) || ischar(rawValue)

        strVal = upper(strtrim(string(rawValue)));

        % Known non-numeric tokens
        if strlength(strVal) == 0 || any(strcmp(strVal, ...
                ["NA","N/A","NAN","NULL","NONE","-","--","MISSING"]))
            numericValue = NaN;
            return;
        end

        % Try numeric conversion
        tmp = str2double(strVal);
        if isnan(tmp)
            numericValue = NaN;
        else
            numericValue = tmp;
        end
        return;
    end

    %------------------------------------------------------
    % Case 6: cell (unwrap scalar cell)
    %------------------------------------------------------
    if iscell(rawValue) && isscalar(rawValue)
        numericValue = sanitizeKpiValue(rawValue{1});
        return;
    end

    %------------------------------------------------------
    % Fallback: any other type -> NaN
    %------------------------------------------------------
    numericValue = NaN;
end
