let    
    bmpPath = "BMP画像のパス",
    Filter   = 
            {
                0, 1, 0,
                1,-4, 1,
                0, 1, 0
            },
    RawTable = Convolution(bmpReader(bmpPath,0),Filter),
    Result   = DispImage(RawTable)
in
    Result