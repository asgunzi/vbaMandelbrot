# vbaMandelbrot

Uma implementação simples dos fractais de Mandelbrot, em Excel - VBA.

Fractais são estruturas auto-similares, belas peças geométricas estudadas por pesquisadores como Benoit Mandelbrot, nos anos 60 e 70.

![](https://ferramentasexcelvba.files.wordpress.com/2018/04/fractals01.jpg)

O conjunto Júlia foi bastante estudado por Mandelbrot, e baseia-se numa equação muito simples:

Z = Z^2 + c

O domínio desta equação é no plano complexo, e indico os seguintes links para entender melhor a teoria.

https://www.wikihow.com/Plot-the-Mandelbrot-Set-By-Hand

https://plus.maths.org/content/computing-mandelbrot-set

A implementação VBA cria um grid de shapes retangulares coloridos, segundo o resultado do conjunto de Mandelbrot.

    ActiveSheet.Shapes.AddShape(msoShapeRectangle, xo, y0, sizeX, sizeY).Select


Exemplos:

Pelo menu do lado, pode-se navegar pela região do espaço, e escolher a resolução (quanto maior, mais demorado para criar).

![](https://ferramentasexcelvba.files.wordpress.com/2018/04/fractals02.jpg)

Detalhamento de outra região do espaço de Mandelbrot.

![](https://ferramentasexcelvba.files.wordpress.com/2018/04/fractals03.jpg)
