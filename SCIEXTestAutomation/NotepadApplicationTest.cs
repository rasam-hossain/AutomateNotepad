using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.IO;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;

namespace SCIEXTestAutomation
{
    /// <summary>
    /// Each and every test can run separately
    /// Which means each of the tests are opening notepad application
    /// and killing it at the end of the test
    /// This is can designed to only launch and kill notepad at the begininng 
    /// and the end of all the tests or a group of tests
    /// For simplicity - each test is prefixed so that they run sequentially to avoid thread locks
    /// </summary>
    /// 
    [TestClass]
    public class NotepadApplicationTest
    {
        // Initialization of the drivers
        WindowsDriver<WindowsElement> notepadSession;
        AppiumOptions desiredCapabilities = new AppiumOptions();


        [TestInitialize()]
        /// <summary>
        /// Start the notepad application using the driver
        /// </summary>
        /// <returns></returns>
        public void StartNotePadApplication()
        {
            // Set the application parameters
            desiredCapabilities.AddAdditionalCapability("app", @"C:\WINDOWS\system32\notepad.exe");
            notepadSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);
        }


        [TestCleanup()]
        /// <summary>
        /// Find all notepad processes in the system and kill them
        /// </summary>
        public void KillAllNotepad()
        {
            Process[] pname = Process.GetProcessesByName("notepad");
            if (pname.Length > 0)
            {
                foreach (var item in pname)
                    item.Kill();
            }
        }

        /// <summary>
        /// Verify that the notepad application can be launched.
        /// </summary>
        [TestMethod]
        public void Test01_VerifyNotepadLaunch()
        {
            /* Steps to the test:
             * 1. TestInitialize() and TestCleanup() will automatically
             * start the notepad application and close it ater its done
             * 2. Verify if its running from the processes
             */

            // Verify if notepad is launched
            Process[] pname = Process.GetProcessesByName("notepad");
            Assert.AreEqual(1, pname.Length);
        }

        /// <summary>
        /// Verify that the notepad editor should open with its default size.
        /// </summary>
        [TestMethod]
        public void Test02_VerifyNotepadDefaultSize()
        {
            /* Steps to the test:
             * 1. TestInitialize() and TestCleanup() will automatically
             * start the notepad application and close it ater its done
             * 2. Verify if the default width = 750 and height = 564
             */
            
            var height = notepadSession.FindElementByName("Untitled - Notepad").Size.Height;
            var width = notepadSession.FindElementByName("Untitled - Notepad").Size.Width;

            Assert.AreEqual(564, height);
            Assert.AreEqual(750, width);
        }

        /// <summary>
        /// Verify that users can write/type alphabets and numeric from a standard keyboard.
        /// </summary>
        [TestMethod]
        public void Test03_VerifyNotepadAcceptsAlphaNumeric()
        {
            /* Steps to the test :
             * 1.TestInitialize() and TestCleanup() will automatically
             * start the notepad application and close it ater its done
             * 2. Write on the notepad editor
             * 3. Verify the data on the editor
             */

            /*
             * This test can be further extended for other type of characters and different variation
             * of the text input.
             */

            // Verify alphabets
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("qwertyuiopasdfghjklzxcvbnm");
            Assert.AreEqual(editWindow.Text, "qwertyuiopasdfghjklzxcvbnm");
            editWindow.Clear();

            // Verify numeric characters
            editWindow.SendKeys("1234567890");
            Assert.AreEqual(editWindow.Text, "1234567890");
            editWindow.Clear();

            // Verify alpha-numeric characters
            editWindow.SendKeys("qwertyuiopasdfghjklzxcvbnm1234567890");
            Assert.AreEqual(editWindow.Text, "qwertyuiopasdfghjklzxcvbnm1234567890");
        }


        /// <summary>
        /// Verify that the user can save the text in a file.
        /// </summary>
        [TestMethod]
        public void Test04_VerifyUserCanSaveNotepadTextInAFile()
        {
            /* Steps to the test :
             * 1. TestInitialize() and TestCleanup() will automatically
             * start the notepad application and close it ater its done
             * 2. Open the notpad application by accessing it's location
             * 3. Edit some text into the notepad edit box and save it
             * 4. Verify if the file exists in the correct location
             */

            // Add texts and click the save button
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("qwertyuiopasdfghjklzxcvbnm");
            notepadSession.FindElementByName("File").Click();
            notepadSession.FindElementByName("Save	Ctrl+S").Click();

            // Save the file as text file
            notepadSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

            // Save it to current directory appending the current DateTime
            string filename = GenerateFileName();
            notepadSession.FindElementByAccessibilityId("1001").SendKeys(filename);
            notepadSession.FindElementByName("Save").Click();


            // Verify if the file has been saved successfully in its location
            // Ignoring the file contents in this test
            while (!File.Exists(filename))
                Thread.Sleep(100);
            bool doesFileExist = File.Exists(@filename);
            Assert.IsTrue(doesFileExist);
        }

        /// <summary>
        /// Verify that the user can open any existing file in notepad.
        /// </summary>
        [TestMethod]
        public void Test05_VerifyUserCanOpenExistingFileInNotepad()
        {
            /* Steps to the test :
             * 1. Read all the text files from current directory
             * 2. Open one of the text file using notepad's UI controls
             * 3. And then compare notepad's edit window with the
             * text file contents using File class
             */

            // Read all .txt files from current directory
            string[] filePaths = Directory.GetFiles(@System.IO.Directory.GetCurrentDirectory(), "*.txt");

            // Open the file from notepad's menu
            notepadSession.FindElementByName("File").Click();
            notepadSession.FindElementByAccessibilityId("2").Click();

            //Wait until the next context menu arrives
            notepadSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

            // Open a random text file
            Random rnd = new Random();
            int r = rnd.Next(filePaths.Length);
            notepadSession.FindElementByAccessibilityId("1148").Click();
            notepadSession.FindElementByAccessibilityId("1148").SendKeys(filePaths[r]);
            notepadSession.FindElementByAccessibilityId("1").Click();


            // Read the text file using the File class and compare it with the edit window on notepad
            string text = System.IO.File.ReadAllText(@filePaths[r]);
            string editWindow = notepadSession.FindElementByClassName("Edit").Text;
            Assert.AreEqual(text, editWindow);
        }

        /// <summary>
        /// Verify that the user can append text to any file and again save the file.
        /// </summary>
        [TestMethod]
        public void Test06_VerifyUserCanAppendExistingFileInNotepad()
        {
            /* Steps to the test :
             * 1. Read all the text files from current directory
             * 2. Open one of the text file using notepad's UI controls
             * 3. And then append on the Edit Window
             * 4. Verify the appended text
             */

            // Read all .txt files from current directory
            string[] filePaths = Directory.GetFiles(@System.IO.Directory.GetCurrentDirectory(), "*.txt");

            // Open the file from notepad's menu
            notepadSession.FindElementByName("File").Click();
            notepadSession.FindElementByAccessibilityId("2").Click();

            //Wait until the next context menu arrives
            notepadSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

            // Open a random text file
            Random rnd = new Random();
            int r = rnd.Next(filePaths.Length);
            notepadSession.FindElementByAccessibilityId("1148").Click();
            notepadSession.FindElementByAccessibilityId("1148").SendKeys(filePaths[r]);
            notepadSession.FindElementByAccessibilityId("1").Click();

            // Append on the Edit Window
            notepadSession.FindElementByAccessibilityId("15").Click();
            string textToAppend = " appending more text";
            notepadSession.FindElementByAccessibilityId("15").SendKeys(textToAppend);

            // save the file
            notepadSession.FindElementByName("File").Click();
            notepadSession.FindElementByName("Save	Ctrl+S").Click();

            // Open the file and verify the text
            string text = System.IO.File.ReadAllText(@filePaths[r]);
            Console.WriteLine(text);
            if (!text.Contains(textToAppend))
                Assert.Fail();
        }

        /// <summary>
        /// Verify that the user can select and delete a text
        /// </summary>
        [TestMethod]
        public void Test07_VerifyUserCanSelectAndDeleteText()
        {
            /* Steps to the test :
             * 1. Add some texts on the notepad editor
             * 2. Select a substring of text from the text and delete it
             * 3. Verify the deletion on the notepad edit window
             */

            // Create a new text file on notepad
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("qwertyuiopasdfghjklzxcvbnm");

            // Go to the beginning of input and select texts until a particular index
            Actions actions = new Actions(notepadSession);
            actions.SendKeys(Keys.Home).Build().Perform();
            string textOnNotepad = notepadSession.FindElementByClassName("Edit").Text;
            int textLength = textOnNotepad.Substring(0, textOnNotepad.IndexOf("a")).Length;
            actions.KeyDown(Keys.LeftShift);
            for (int i = 0; i < textLength; i++)
            {
                actions.SendKeys(Keys.ArrowRight);
            }
            actions.KeyUp(Keys.LeftShift);
            actions.Build().Perform();
            editWindow.SendKeys(Keys.Delete);
            Console.WriteLine();

            // Verification of the text in the edit window
            Assert.AreEqual("asdfghjklzxcvbnm", editWindow.Text);
        }

        /// <summary>
        /// Verify that the user can undo any latest change done in the file.
        /// </summary>
        [TestMethod]
        public void Test08_VerifyUserCanUndoLatestChange()
        {
            /* Steps to the test :
             * 1. Add some texts on the notepad editor
             * 2. Use notepad's UI button to undo and verify the undo operation
             * 3. Use windows shortcut key to undo and verify the undo operation 
             */

            // Add texts and click the save button
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("Verify if undo option is working - Method 1");

            // Method - 1 : using the notepad UI control
            notepadSession.FindElementByName("Edit").Click();
            notepadSession.FindElementByAccessibilityId("16").Click();
            Assert.AreEqual("", editWindow.Text);

            // Method - 2 : using the keyboard shortcut (ctrl + z)
            editWindow.SendKeys("Verify if undo option is working - Mehtod 2");
            Actions actions = new Actions(notepadSession);
            actions.KeyDown(Keys.Control).SendKeys("z").KeyUp(Keys.Control).Perform();
            Assert.AreEqual("Verify if undo option is working - Method 1", editWindow.Text);
        }

        /// <summary>
        /// Verify that the user can redo any latest change done in the file.
        /// </summary>
        [TestMethod]
        public void Test09_VerifyUserCanRedoLatestChange()
        {

            /* Steps to the test :
             * 1. Add some texts on the notepad edit window
             * 2. First undo and then redo the texts
             * 3. Verify the redo operation
             */

            // Add texts and click the save button
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("Verify if redo option is working");

            // Method - 1 : using the notepad UI control
            notepadSession.FindElementByName("Edit").Click();
            notepadSession.FindElementByAccessibilityId("16").Click();

            /*
             * Redo shortbut : ctrl + shift + z
             */
            // First Undo and then Redo the changes
            Actions actions = new Actions(notepadSession);
            actions.KeyDown(Keys.Control).KeyDown(Keys.Shift).SendKeys("z").KeyUp(Keys.Control).KeyUp(Keys.Shift).Perform();
            Assert.AreEqual("Verify if redo option is working", editWindow.Text);
        }

        /// <summary>
        /// Verify that the user can close the editor window by clicking the cross icon.
        /// </summary>
        [TestMethod]
        public void Test10_VerifyUserCloseNotepad()
        {

            /* Steps to the test :
             * 0. Some actions are performed before closing the notepad
             * 1. Add some texts on the notepad edit window and save the file
             * 2. Close notepad
             */

            // Add texts and click the save button
            var editWindow = notepadSession.FindElementByClassName("Edit");
            editWindow.SendKeys("qwertyuiopasdfghjklzxcvbnm");
            notepadSession.FindElementByName("File").Click();
            notepadSession.FindElementByName("Save	Ctrl+S").Click();

            // Save the file as text file
            notepadSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

            // Save it to current directory appenidng the current DateTime
            string filename = GenerateFileName();
            notepadSession.FindElementByAccessibilityId("1001").SendKeys(filename);
            notepadSession.FindElementByName("Save").Click();

            // Close the notepad
            notepadSession.FindElementByName("Close").Click();

            // Verify if notepad is closed properly
            Process[] pname = Process.GetProcessesByName("notepad");
            Console.WriteLine(pname.Length);
            Assert.AreEqual(0, pname.Length);
        }

        /// <summary>
        /// Generates a filename based on current datetime
        /// </summary>
        /// <returns></returns>
        public string GenerateFileName()
        {
            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            currentDirectory.Replace(@"\", @"\\");
            var currentDateTime = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss-fff");
            string filename = $"{currentDirectory}\\{currentDateTime}.txt";
            return filename;
        }
    }
}
