﻿using System;
using FileFormat.Slides.Examples;

namespace FileFormat.Slides.Examples.Usage
{
    class Program
    {
        static void Main (string[ ] args)
        {
            // SlideExamples slideExamples = new SlideExamples();
            // slideExamples.CreateNewSlideInNewPresentation();

            // slideExamples.CreateNewSlideInExistingPresentation(filename:"test.pptx");

            // slideExamples.RemoveSlideInAnExistingPresentation(filename: "test.pptx");

            //TextExamples textExamples = new TextExamples();
            //textExamples.CreateNewTextShapeInNewSlide();
            //textExamples.AddNewTextShapeExistingSlide(filename: "sample.pptx");

            ImageExamples imageExamples = new ImageExamples();
            //imageExamples.AddImageInASlide(imagename:"sample.jpg");
            imageExamples.UpdateImageInExistingSlide(filename: "sample.pptx", xAxis: 300.0, yAxis: 200.0);

        }
    }
}