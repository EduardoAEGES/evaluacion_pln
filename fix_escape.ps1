$file = 'c:\Users\Usuario\Desktop\FORMULARIO_PLN\admin.html'
$content = [System.IO.File]::ReadAllText($file)
# Replace all \\\\  (4 backslashes in file = \\  in regex) with \\  (2 backslashes = \ in output)
# We need to fix: \\alpha -> \alpha, etc in the EQ_SYMBOLS block
$content = $content -replace '\\\\\\\\(times|div|neq|leq|geq|pm|mp|cdot|left|right|ldots|infty|approx|propto|circ|frac|sqrt|partial|overline|hat|tilde|bar|vec|dot|alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Sigma|Phi|Psi|Omega|int|iint|oint|sum|prod|lim|to|nabla|ln|log|sin|cos|tan|in|notin|subset|supset|subseteq|supseteq|cup|cap|emptyset|mathbb|forall|exists|land|lor|neg|rightarrow|leftarrow|leftrightarrow|Rightarrow|Leftarrow|Leftrightarrow|uparrow|downarrow|mapsto|implies|impliedby|therefore|because)', '\$1'
[System.IO.File]::WriteAllText($file, $content)
Write-Host "Done"
