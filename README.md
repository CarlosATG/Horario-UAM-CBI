GPT

¡Genial! Me alegra mucho saber que ahora funciona correctamente. Aquí te dejo una breve explicación del programa en español, incluyendo las restricciones implementadas y el resultado esperado:
Descripción del Programa

Este programa genera un horario académico para varias materias, profesores, grupos y salones, tomando en cuenta diversas restricciones y condiciones. La aplicación cuenta con una interfaz gráfica para facilitar la asignación de los profesores y las frecuencias de clases por grupo. Luego, el programa genera un horario que respeta las restricciones y lo exporta a diferentes archivos Excel.
Restricciones Implementadas

    Número de clases por grupo:
        Un grupo puede tener hasta 4 clases al día, con la condición de que no se programen más de 2 clases consecutivas sin un descanso de 1 hora entre ellas.

    Restricción para los profesores:
        Un profesor no puede tener más de 2 clases consecutivas y un máximo de 3 clases al día.

    Días consecutivos:
        Una clase no puede ser programada en 3 días consecutivos. Si un grupo ya tiene clases en dos días consecutivos, el tercer día será saltado para esa clase.

    Duración de las clases:
        Existen tres tipos de duraciones: 90 minutos, 120 minutos y 180 minutos. Estas clases se programan con la frecuencia adecuada (ej.: 90 minutos, 3 veces a la semana; 120 minutos o 180 minutos, una vez a la semana).

Salida del Programa

El programa genera tres archivos Excel con la siguiente información:

    Horario por grupo: Cada hoja de este archivo representa el horario de un grupo, mostrando las clases asignadas, los profesores, los salones y los horarios de inicio y fin.
    Horario por profesor: En este archivo, cada hoja representa el horario de un profesor, mostrando las clases que imparte, los grupos correspondientes y el aula donde enseña.
    Horario por salón: Cada hoja en este archivo corresponde a un aula, mostrando las clases programadas para ese salón, el grupo y el profesor que ocupa esa aula en cada bloque de tiempo.

Resultado Esperado

El resultado es un conjunto de horarios que respetan las restricciones de clases consecutivas, descanso para los grupos, profesores y la correcta distribución de las clases en el calendario semanal. Los horarios se generan de manera ordenada y lógica, asegurando que no haya conflictos de tiempo entre los grupos, profesores y aulas.
