let
    h     = 0.01,
    steps = 10000,
    Langford = (x as list) as list =>
    let
        x1 = x{0},
        x2 = x{1},
        x3 = x{2},
        alpha = 1.0,
        beta  = 0.7,
        gamma = 0.6,
        omega = 3.5,
        rho   = 0.25,
        eps   = 0.0,
        dx1 = (x3 - beta) * x1 - omega * x2,
        dx2 = omega * x1 + (x3 - beta) * x2,
        dx3 = gamma + alpha * x3 - (x3*x3*x3)/3.0 - (x1*x1 + x2*x2) * (1 + rho * x3) + eps * x3 * (x1*x1*x1)
    in
        {dx1, dx2, dx3},
    result = rk4_function(Langford, h, steps, {0.1, 0.0, 0.0})
in
    result