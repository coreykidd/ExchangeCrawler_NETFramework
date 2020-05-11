using System.Security;

namespace Authentication
{
    interface ICredentialsInterface
    {
        string EmailAddress { get; set; }
        SecureString Password { get; set; }
    }
}

//public class Credentials : ICredentialsInterface
//{
//    public string EmailAddress { get; set; }
//    public SecureString Password { get; set; }

//    public Credentials()
//    {
//        this.EmailAddress = "testuser";
//        this.Password = new SecureString();
//        this.Password.AppendChar('p');
//        this.Password.AppendChar('a');
//        this.Password.AppendChar('s');
//        this.Password.AppendChar('s');
//        this.Password.AppendChar('w');
//        this.Password.AppendChar('o');
//        this.Password.AppendChar('r');
//        this.Password.AppendChar('d');
//        this.Password.MakeReadOnly();
//    }
//}