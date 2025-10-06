let
    h     = 0.01,
    steps = 10000,
    Rossler = (x as list) as list =>
    let
        x1 = x{0},
        x2 = x{1},
        x3 = x{2},

        a = 0.2,
        b = 0.2,
        c = 5.7,

        dx1 = -x2 - x3,
        dx2 = x1 + a * x2,
        dx3 = b + x3 * (x1 - c)
    in
        {dx1, dx2, dx3},
    result = rk4_function(Rossler, h, steps, {0.1, 0.0, 0.0})
in
    result