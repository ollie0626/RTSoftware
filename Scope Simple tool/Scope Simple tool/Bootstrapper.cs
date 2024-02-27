﻿using System;
using Stylet;
using StyletIoC;
using Scope_Simple_tool.Pages;

namespace Scope_Simple_tool
{
    public class Bootstrapper : Bootstrapper<ShellViewModel>
    {
        protected override void ConfigureIoC(IStyletIoCBuilder builder)
        {
            // Configure the IoC container in here
        }

        protected override void Configure()
        {
            // Perform any other configuration before the application starts
        }
    }
}
