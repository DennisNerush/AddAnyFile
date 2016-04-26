# VS Test Generator

Download the extension at the .....

-------------------------------------------------

A Visual Studio extension for easily creating test class for the current working file.

See the [changelog](CHANGELOG.md) for updates and roadmap.

### Features

- Easily create test file for the current class
- Set the current class as the target of the test
- Add Nunit TestFixture attribute

### Show the dialog

A new button is added to the context menu in Solution Explorer.

![Add new file dialog](art/menu.png)

You can either click that button or use the keybord shortcut **Shift+F2**.

### Output
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using NUnit.Framework;
    
    namespace TestNamespace
    {
    	[TestFixture]
        class YourClassTests : AbstractTestFixture
        {
    		private YourClass _target;
    
    		public YourClassTests() 
    		{
    			_target = new YourClass();
    		}
    
    		[Test]
    		public void Method_When_Should()
    		{
    			//arrange
    
    			//act
    
    			//assert
    		}
        }
    }


## Contribute
Check out the [contribution guidelines](.github/CONTRIBUTING.md)
if you want to contribute to this project.

For cloning and building this project yourself, make sure
to install the
[Extensibility Tools 2015](https://visualstudiogallery.msdn.microsoft.com/ab39a092-1343-46e2-b0f1-6a3f91155aa6)
extension for Visual Studio which enables some features
used by this project.

## License
[Apache 2.0](LICENSE)