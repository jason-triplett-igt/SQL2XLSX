using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace SQLScript2XLSX_2.Tests
{
    [TestClass]
    public class PasswordBoxHelperTests
    {
        [TestMethod]
        public void BoundPassword_ShouldUpdatePasswordBox()
        {
            var thread = new Thread(() =>
            {
                // Arrange
                var passwordBox = new PasswordBox();
                PasswordBoxHelper.SetBoundPassword(passwordBox, "initialPassword");

                // Act
                PasswordBoxHelper.SetBoundPassword(passwordBox, "newPassword");

                // Assert
                Assert.AreEqual("newPassword", passwordBox.Password);
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        [TestMethod]
        public void PasswordBox_ShouldUpdateBoundPassword()
        {
            var thread = new Thread(() =>
            {
                // Arrange
                var passwordBox = new PasswordBox();
                PasswordBoxHelper.SetBindPassword(passwordBox, true);
                PasswordBoxHelper.SetBoundPassword(passwordBox, "initialPassword");

                // Act
                passwordBox.Password = "newPassword";

                // Assert
                Assert.AreEqual("newPassword", PasswordBoxHelper.GetBoundPassword(passwordBox));
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        [TestMethod]
        public void BindPassword_ShouldAttachAndDetachEventHandlers()
        {
            var thread = new Thread(() =>
            {
                // Arrange
                var passwordBox = new PasswordBox();
                bool eventHandlerCalled = false;
                RoutedEventHandler handler = (sender, args) => eventHandlerCalled = true;

                // Act
                PasswordBoxHelper.SetBindPassword(passwordBox, true);
                passwordBox.PasswordChanged += handler;
                passwordBox.Password = "test";
                bool isHandlerAttached = eventHandlerCalled;

                eventHandlerCalled = false;
                PasswordBoxHelper.SetBindPassword(passwordBox, false);
                passwordBox.PasswordChanged -= handler;
                passwordBox.Password = "test2";
                bool isHandlerDetached = !eventHandlerCalled;

                // Assert
                Assert.IsTrue(isHandlerAttached);
                Assert.IsTrue(isHandlerDetached);
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }
    }
}
