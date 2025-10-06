let
    h     = 0.01,
    steps = 10000,
    Lorenz = (x as list) as list =>
    let
        sigma = 10.0,
        rho   = 28.0,
        beta  = 8.0/3.0,
        dx = sigma * (x{1} - x{0}),
        dy = x{0} * (rho - x{2}) - x{1},
        dz = x{0} * x{1} - beta * x{2}
    in
        {dx, dy, dz},

    result = rk4_function(Lorenz,h,steps,{1,1,1})
in
    result
