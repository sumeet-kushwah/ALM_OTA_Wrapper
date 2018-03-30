
namespace ALM_Wrapper
{
    public class Releases
    {
        public Release Release;
        public ReleaseFolders ReleaseFolders;
        public Cycle Cycle;

        public Releases(TDAPIOLELib.TDConnection tDConnection)
        {
            Release = new Release(tDConnection);
            ReleaseFolders = new ReleaseFolders(tDConnection);
            Cycle = new Cycle(tDConnection);
        }

    }
}
