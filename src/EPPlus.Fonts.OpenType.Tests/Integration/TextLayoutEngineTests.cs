using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Diagnostics;
using EPPlus.Fonts.OpenType.Utils;

namespace EPPlus.Fonts.OpenType.Tests.Integration
{
    [TestClass]
    public class TextLayoutEngineTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Single-Font Wrapping Tests

        [TestMethod]
        public void WrapText_ShortText_NoWrapping()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper, FontFolders);

            // Act
            var lines = layout.WrapText("Hello", 11f, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello", lines[0]);
        }

        [TestMethod]
        public void WrapText_LongText_WrapsAtSpaces()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act - narrow width forces wrapping
            var lines = layout.WrapText("Hello world test", 11f, 50);

            // Assert
            Assert.IsTrue(lines.Count > 1, "Text should wrap to multiple lines");

            // Each line should be a complete word (no mid-word breaks)
            foreach (var line in lines)
            {
                Assert.IsFalse(string.IsNullOrEmpty(line));
            }
        }

        [TestMethod]
        public void WrapText_WithLineBreaks_PreservesBreaks()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapText("Line 1\r\nLine 2\nLine 3", 11f, 1000);

            // Assert
            Assert.AreEqual(3, lines.Count);
            Assert.AreEqual("Line 1", lines[0]);
            Assert.AreEqual("Line 2", lines[1]);
            Assert.AreEqual("Line 3", lines[2]);
        }

        [TestMethod]
        public void WrapText_TestWhenOnExactWrapPlusSpaces2()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Aptos Narrow", FontSubFamily.Regular);
            //'sit amet non' is EXACTLY 72 pixels (54 points) in excel at 100% size/display
            //So an added space should push 'non' over the edge to the next line
            var text = "sit amet  non lacus.";
            var comparison = new List<string>() {"sit amet", "non lacus."};

            var maxWidthPoints = 54d;

            ITextShaper shaper = new TextShaper(font);
            using var layoutEngine = new TextLayoutEngine(shaper);
            var wrappedLines = layoutEngine.WrapText(
                text,
                11f,
                maxWidthPoints,
                ShapingOptions.Full
            );

            Assert.IsTrue(comparison.SequenceEqual(wrappedLines));
        }

        [TestMethod]
        public void WrapText_TestWhenOnExactWrap()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Aptos Narrow", FontSubFamily.Regular);
            //'sit amet non' is EXACTLY 72 pixels (54 points) in excel at 100% size/display
            var text = "nulla efficitur commodo sit amet non lacus. Proin viverra enim";
            var comparison = new List<string>() { "nulla", "efficitur", "commodo", "sit amet non", "lacus. Proin", "viverra enim" };

            ITextShaper shaper = new TextShaper(font);
            using var layoutEngine = new TextLayoutEngine(shaper);
            var wrappedLines = layoutEngine.WrapText(
                text,
                11f,
                54,
                ShapingOptions.Full
            );

            Assert.IsTrue(comparison.SequenceEqual(wrappedLines));
        }

        [TestMethod]
        public void WrapText_TestFragments()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Aptos Narrow", FontSubFamily.Regular);
            var Lorem20Str = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla pulvinar interdum imperdiet. Praesent ut auctor urna. Phasellus sollicitudin quam vitae est convallis, eu mattis lorem efficitur. Mauris nulla libero, tincidunt id ipsum non, lobortis tristique mauris. Donec ut enim sed enim fermentum molestie vel quis odio. Morbi a fermentum massa, sit amet ultrices est. Aenean ante mi, fermentum nec rhoncus et, vulputate vel sapien. Donec tempus, leo quis luctus rhoncus, augue odio pharetra libero, ac blandit urna turpis sed diam. Vivamus augue purus, eleifend et justo facilisis, imperdiet rhoncus sem. Quisque accumsan pellentesque elit, eget finibus massa accumsan in. Fusce eu accumsan enim. Cras pulvinar enim vel tellus lacinia, consectetur euismod tortor consectetur. Praesent tincidunt pretium eros, ac auctor magna luctus sed. Ut porta lectus quam, non ornare mauris lacinia sit amet. Nullam egestas dolor quis magna porttitor, ac iaculis nisi hendrerit. Proin at mollis lacus, in porttitor nunc. Aliquam erat volutpat. Sed vel egestas risus, at aliquam arcu. Vestibulum quis lobortis nulla. Etiam pellentesque auctor nulla, eget tincidunt felis rhoncus id. Sed metus ante, efficitur id dui eu, fermentum mollis odio. Phasellus ullamcorper iaculis augue vel consequat. Etiam fringilla euismod interdum. Ut molestie massa id fringilla lobortis. Vestibulum malesuada, ante vel mattis ultrices, sem ante molestie augue, non tristique dui mi non nibh. Maecenas dictum, sem eget convallis rhoncus, lacus enim porta neque, in posuere dui ex a sapien. Nam lacus nibh, posuere sed elit eget, condimentum facilisis ligula. Cras consectetur lacus ullamcorper velit aliquet bibendum eget vel nulla. Aenean varius ac erat quis ullamcorper. Donec laoreet arcu a lorem volutpat faucibus. Vivamus vehicula leo ut erat luctus scelerisque. Morbi posuere ex et magna egestas facilisis. Fusce scelerisque volutpat erat bibendum hendrerit. Nam blandit mi ut metus pulvinar, vel tempus lacus euismod. Quisque imperdiet sit amet sapien sed ultricies. Phasellus sodales, ipsum vitae tincidunt facilisis, nulla ligula faucibus felis, eget vehicula ante lacus eu lorem. Integer congue diam ac viverra tristique. Curabitur tristique dolor quis quam pretium, et scelerisque quam dictum. Maecenas vitae sodales ligula. Pellentesque maximus diam vel porta convallis. Ut aliquam eros quis porta pellentesque. Fusce in ex ut mi egestas cursus. Aliquam erat volutpat. Cras laoreet condimentum laoreet. Sed eget facilisis tellus. Morbi viverra odio sed odio placerat mollis. Duis turpis metus, dignissim varius urna quis, viverra dignissim dui. Vivamus viverra at nisi quis convallis. Suspendisse fringilla risus et ante sollicitudin, sed eleifend sem placerat. Proin pretium blandit arcu, eget rhoncus risus hendrerit at. Interdum et malesuada fames ac ante ipsum primis in faucibus. Phasellus vulputate efficitur maximus. Cras blandit nulla eu nisi auctor tempus. Sed pretium lacus ac magna vestibulum, aliquam faucibus orci luctus. Mauris enim lorem, varius ut ante quis, varius viverra lectus. Fusce blandit nibh vel feugiat efficitur. Donec maximus id justo ac mollis. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla placerat lectus et purus dictum, id congue nisi euismod. Maecenas euismod fermentum diam, sit amet gravida magna suscipit a. Quisque consectetur arcu eu nunc sodales scelerisque. Nulla non tincidunt nulla. Pellentesque ut tortor vel enim convallis malesuada. Aliquam ultricies bibendum ultrices. Mauris rutrum ac nisl vel luctus. Donec quis nibh vitae orci ultricies gravida. Aliquam vitae velit porttitor lorem bibendum fringilla volutpat a eros. Curabitur at commodo tortor. Etiam ultricies, neque et iaculis euismod, diam ligula luctus mi, vitae lobortis felis lorem eu nulla. Sed a semper ex. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nulla mauris elit, pulvinar ac tortor et, luctus hendrerit nisl. In egestas auctor urna vitae laoreet. Praesent bibendum egestas convallis. Proin non suscipit tellus. Nullam at nibh in urna laoreet sodales non vel tellus. Donec in enim dui. Phasellus quis quam tincidunt, pellentesque lorem ac, scelerisque neque. Integer nec tempus urna. Donec elit massa, eleifend eu sapien sit amet, mollis pellentesque est. Nullam tristique tellus iaculis arcu consectetur pretium. Sed venenatis convallis scelerisque. Suspendisse varius urna sit amet purus accumsan, id ultricies erat efficitur. Cras non ipsum eget nulla efficitur commodo sit amet non lacus. Proin viverra enim sit amet enim tempus ullamcorper. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Duis ac massa interdum, gravida ex egestas, finibus purus. Nunc consectetur commodo lacus, ac convallis quam lobortis eu. Sed convallis tempor commodo. Nulla sed convallis mauris. Donec venenatis nisi est, ac ullamcorper mi pretium quis. Donec vitae eros at ipsum interdum scelerisque nec vitae nisi. Sed vestibulum erat ac bibendum dapibus. Morbi nec elit id quam tristique cursus id sed sem. Praesent non ante enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Praesent non mauris dui. Aliquam rhoncus mattis ante sed venenatis. Vivamus vehicula sed sapien sed dictum. In aliquet, urna efficitur tincidunt lobortis, nibh justo tristique purus, sed volutpat risus magna et libero.Suspendisse lectus justo, varius eget arcu et, semper laoreet erat. Quisque eget lacus ornare, pellentesque erat sit amet, vulputate felis. Duis luctus, massa a pellentesque mollis, massa elit convallis mi, vel bibendum ex ex eu purus. Suspendisse vel fermentum urna, ac commodo enim. Mauris tincidunt cursus elit, a volutpat libero commodo et. Etiam dapibus libero venenatis tellus lobortis, vel lacinia elit faucibus. Maecenas semper sed quam quis finibus. Integer efficitur, libero imperdiet sollicitudin commodo, elit arcu vulputate est, eget finibus mi urna sit amet magna. Cras ullamcorper consequat ornare. Fusce convallis nunc vel risus cursus, at maximus ligula cursus. Pellentesque vulputate risus libero, eget cursus nibh sodales sed. Donec accumsan sem et massa semper, id dignissim velit vehicula.Cras cursus ipsum ac erat vehicula, nec iaculis purus dictum. Quisque lacinia elit vitae leo dictum, vel dignissim velit dapibus. Aenean sem nisi, faucibus interdum justo eu, euismod porttitor ex. Morbi et lectus lectus. Duis neque felis, suscipit at scelerisque eu, scelerisque id orci. Curabitur et placerat ipsum. Proin gravida sapien nisl, et varius ipsum mollis nec. Quisque dignissim consectetur feugiat. Aenean eros purus, laoreet interdum rutrum at, aliquet sit amet lectus. Donec gravida lorem ut tincidunt laoreet. Donec consequat viverra ligula, in accumsan mi bibendum scelerisque. Quisque ac risus justo. Morbi magna arcu, egestas nec luctus commodo, cursus eget nunc. Vivamus euismod lorem ex, et maximus felis hendrerit eget. Nullam ullamcorper euismod ligula, et iaculis ligula ultricies a. Fusce aliquam, enim vel fermentum ultrices, elit quam semper erat, vitae semper velit augue non magna.Quisque maximus semper arcu, id pellentesque est tempus a. Phasellus lacus elit, auctor sit amet lacinia a, dapibus vitae velit. Phasellus ut pharetra justo, ut ultricies erat. Sed molestie sapien vel interdum lobortis. Nulla facilisi. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla nec mauris quis nisi vulputate gravida quis nec velit.Nam et congue ipsum. Nulla vel elit non dolor mollis aliquet vel at magna. Pellentesque nec facilisis elit. In vulputate quis sem porta suscipit. Nullam sed ex ornare nibh suscipit mattis quis non lacus. Mauris vel ex urna. Vivamus ultricies sapien sit amet sapien vehicula gravida. Donec feugiat volutpat quam. Vestibulum auctor dictum nisl, id hendrerit metus ullamcorper sed. Nulla maximus lacus vel mollis maximus. Nulla laoreet placerat quam eu viverra. Etiam feugiat accumsan nisl a condimentum. Sed ultricies ante ante, ac auctor ligula gravida nec. Praesent a neque dignissim, sagittis felis sit amet, condimentum turpis. Fusce at leo vel est blandit malesuada. Pellentesque et neque non metus pellentesque imperdiet. Praesent pellentesque lacinia lorem, et tristique tellus efficitur id. Suspendisse aliquet ultricies justo vitae interdum. Cras tristique viverra quam, eget gravida mi fermentum imperdiet. Sed imperdiet vitae purus ut volutpat. Nulla lacinia elit in fermentum consectetur. Phasellus commodo ut nisl sit amet sagittis. Duis ac ornare orci. Vivamus vel enim posuere, pharetra ex vel, elementum est. Vestibulum commodo luctus metus eget maximus. Suspendisse a nulla a odio eleifend faucibus. Suspendisse semper lacus non porttitor aliquet. Cras ac scelerisque magna, et pulvinar justo. Integer cursus pulvinar fringilla. Mauris imperdiet nibh sit amet tempor laoreet. Morbi tincidunt tortor ex, sit amet maximus purus tristique quis. Quisque sed hendrerit velit. Mauris mattis nibh ut eros luctus, eget mattis massa auctor. Phasellus eu neque at augue gravida sagittis nec non tortor. Etiam porttitor sem sodales mi ullamcorper gravida. In in dictum orci. In vitae vestibulum quam. Cras augue eros, tincidunt ac elit posuere, sollicitudin efficitur lectus. Praesent quis sodales nisl. Proin sit amet molestie est. In commodo mauris vel mauris efficitur, nec mollis mauris sagittis. Cras ligula nibh, egestas sit amet eros in, lacinia tristique magna. Cras risus libero, lacinia eget libero vitae, maximus aliquet nibh. Mauris id sodales purus, vitae dictum lectus. Cras consectetur ligula velit, tempus pulvinar lacus porttitor vitae. Phasellus eget tellus ipsum. Donec interdum laoreet elit non vestibulum. Cras sed urna ullamcorper, aliquam erat eget, porta orci. Vestibulum eget congue nulla. Sed sem tortor, euismod at rutrum id, sagittis a nunc. Duis in nibh facilisis, dignissim purus ut, hendrerit magna. Sed semper ligula id massa elementum, non malesuada velit egestas. Nullam dictum, mi nec euismod sagittis, ligula leo ullamcorper dolor, quis faucibus odio metus eget magna. Ut gravida metus non metus bibendum bibendum. In sagittis eleifend aliquet. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nam mollis sagittis felis, in faucibus tortor pretium vel. Nam nec enim metus. Donec in augue arcu. Proin non lobortis purus, sit amet lacinia elit. Suspendisse quis eros condimentum, blandit justo sit amet, lobortis nisl. Suspendisse maximus massa sed urna tempor ornare. Nunc malesuada purus odio, eu luctus lectus auctor nec. Morbi auctor pellentesque auctor. Sed ullamcorper, ex vitae aliquam vulputate, est diam feugiat mi, id porttitor lectus orci ac leo. Donec sit amet velit pulvinar, venenatis turpis ut, interdum ligula. Interdum et malesuada fames ac ante ipsum primis in faucibus. Vestibulum eu lacus urna. Maecenas sem nulla, accumsan eu ultricies sed, tempor vel magna. Cras aliquet sollicitudin sapien ac pulvinar. Praesent ac sodales mi. Integer vitae mauris massa. Maecenas iaculis orci et faucibus interdum. Nunc nec maximus felis, sed finibus quam. Pellentesque felis massa, vestibulum in tellus vitae, congue tincidunt justo. Nunc vitae enim malesuada, bibendum ante nec, varius tellus. Praesent vitae nisi id quam auctor lacinia at non quam. Nam nec ligula sit amet felis auctor sagittis. Nunc in risus eu urna varius laoreet quis sit amet felis. Morbi varius tempor orci, eu vestibulum nunc vestibulum ac. Nunc vehicula velit eleifend consequat porta. Suspendisse maximus dapibus orci, in vulputate massa pretium ac. Quisque malesuada aliquet aliquet.";

            const string SavedComparisonString = "Lorem\r\nipsum dolor\r\nsit amet,\r\nconsectetur\r\nadipiscing\r\nelit. Nulla\r\npulvinar\r\ninterdum\r\nimperdiet.\r\nPraesent ut\r\nauctor urna.\r\nPhasellus\r\nsollicitudin\r\nquam vitae\r\nest\r\nconvallis,\r\neu mattis\r\nlorem\r\nefficitur.\r\nMauris nulla\r\nlibero,\r\ntincidunt id\r\nipsum non,\r\nlobortis\r\ntristique\r\nmauris.\r\nDonec ut\r\nenim sed\r\nenim\r\nfermentum\r\nmolestie vel\r\nquis odio.\r\nMorbi a\r\nfermentum\r\nmassa, sit\r\namet\r\nultrices est.\r\nAenean\r\nante mi,\r\nfermentum\r\nnec\r\nrhoncus et,\r\nvulputate\r\nvel sapien.\r\nDonec\r\ntempus, leo\r\nquis luctus\r\nrhoncus,\r\naugue odio\r\npharetra\r\nlibero, ac\r\nblandit urna\r\nturpis sed\r\ndiam.\r\nVivamus\r\naugue\r\npurus,\r\neleifend et\r\njusto\r\nfacilisis,\r\nimperdiet\r\nrhoncus\r\nsem.\r\nQuisque\r\naccumsan\r\npellentesqu\r\ne elit, eget\r\nfinibus\r\nmassa\r\naccumsan\r\nin. Fusce eu\r\naccumsan\r\nenim. Cras\r\npulvinar\r\nenim vel\r\ntellus\r\nlacinia,\r\nconsectetur\r\neuismod\r\ntortor\r\nconsectetur\r\n. Praesent\r\ntincidunt\r\npretium\r\neros, ac\r\nauctor\r\nmagna\r\nluctus sed.\r\nUt porta\r\nlectus\r\nquam, non\r\nornare\r\nmauris\r\nlacinia sit\r\namet.\r\nNullam\r\negestas\r\ndolor quis\r\nmagna\r\nporttitor, ac\r\niaculis nisi\r\nhendrerit.\r\nProin at\r\nmollis\r\nlacus, in\r\nporttitor\r\nnunc.\r\nAliquam\r\nerat\r\nvolutpat.\r\nSed vel\r\negestas\r\nrisus, at\r\naliquam\r\narcu.\r\nVestibulum\r\nquis\r\nlobortis\r\nnulla. Etiam\r\npellentesqu\r\ne auctor\r\nnulla, eget\r\ntincidunt\r\nfelis\r\nrhoncus id.\r\nSed metus\r\nante,\r\nefficitur id\r\ndui eu,\r\nfermentum\r\nmollis odio.\r\nPhasellus\r\nullamcorper\r\niaculis\r\naugue vel\r\nconsequat.\r\nEtiam\r\nfringilla\r\neuismod\r\ninterdum.\r\nUt molestie\r\nmassa id\r\nfringilla\r\nlobortis.\r\nVestibulum\r\nmalesuada,\r\nante vel\r\nmattis\r\nultrices,\r\nsem ante\r\nmolestie\r\naugue, non\r\ntristique dui\r\nmi non\r\nnibh.\r\nMaecenas\r\ndictum,\r\nsem eget\r\nconvallis\r\nrhoncus,\r\nlacus enim\r\nporta\r\nneque, in\r\nposuere dui\r\nex a sapien.\r\nNam lacus\r\nnibh,\r\nposuere sed\r\nelit eget,\r\ncondimentu\r\nm facilisis\r\nligula. Cras\r\nconsectetur\r\nlacus\r\nullamcorper\r\nvelit aliquet\r\nbibendum\r\neget vel\r\nnulla.\r\nAenean\r\nvarius ac\r\nerat quis\r\nullamcorper\r\n. Donec\r\nlaoreet arcu\r\na lorem\r\nvolutpat\r\nfaucibus.\r\nVivamus\r\nvehicula leo\r\nut erat\r\nluctus\r\nscelerisque.\r\nMorbi\r\nposuere ex\r\net magna\r\negestas\r\nfacilisis.\r\nFusce\r\nscelerisque\r\nvolutpat\r\nerat\r\nbibendum\r\nhendrerit.\r\nNam blandit\r\nmi ut metus\r\npulvinar, vel\r\ntempus\r\nlacus\r\neuismod.\r\nQuisque\r\nimperdiet\r\nsit amet\r\nsapien sed\r\nultricies.\r\nPhasellus\r\nsodales,\r\nipsum vitae\r\ntincidunt\r\nfacilisis,\r\nnulla ligula\r\nfaucibus\r\nfelis, eget\r\nvehicula\r\nante lacus\r\neu lorem.\r\nInteger\r\ncongue\r\ndiam ac\r\nviverra\r\ntristique.\r\nCurabitur\r\ntristique\r\ndolor quis\r\nquam\r\npretium, et\r\nscelerisque\r\nquam\r\ndictum.\r\nMaecenas\r\nvitae\r\nsodales\r\nligula.\r\nPellentesqu\r\ne maximus\r\ndiam vel\r\nporta\r\nconvallis. Ut\r\naliquam\r\neros quis\r\nporta\r\npellentesqu\r\ne. Fusce in\r\nex ut mi\r\negestas\r\ncursus.\r\nAliquam\r\nerat\r\nvolutpat.\r\nCras laoreet\r\ncondimentu\r\nm laoreet.\r\nSed eget\r\nfacilisis\r\ntellus.\r\nMorbi\r\nviverra odio\r\nsed odio\r\nplacerat\r\nmollis. Duis\r\nturpis\r\nmetus,\r\ndignissim\r\nvarius urna\r\nquis, viverra\r\ndignissim\r\ndui.\r\nVivamus\r\nviverra at\r\nnisi quis\r\nconvallis.\r\nSuspendiss\r\ne fringilla\r\nrisus et ante\r\nsollicitudin,\r\nsed eleifend\r\nsem\r\nplacerat.\r\nProin\r\npretium\r\nblandit\r\narcu, eget\r\nrhoncus\r\nrisus\r\nhendrerit at.\r\nInterdum et\r\nmalesuada\r\nfames ac\r\nante ipsum\r\nprimis in\r\nfaucibus.\r\nPhasellus\r\nvulputate\r\nefficitur\r\nmaximus.\r\nCras blandit\r\nnulla eu nisi\r\nauctor\r\ntempus.\r\nSed pretium\r\nlacus ac\r\nmagna\r\nvestibulum,\r\naliquam\r\nfaucibus\r\norci luctus.\r\nMauris enim\r\nlorem,\r\nvarius ut\r\nante quis,\r\nvarius\r\nviverra\r\nlectus.\r\nFusce\r\nblandit nibh\r\nvel feugiat\r\nefficitur.\r\nDonec\r\nmaximus id\r\njusto ac\r\nmollis.\r\nVestibulum\r\nante ipsum\r\nprimis in\r\nfaucibus\r\norci luctus\r\net ultrices\r\nposuere\r\ncubilia\r\ncurae; Nulla\r\nplacerat\r\nlectus et\r\npurus\r\ndictum, id\r\ncongue nisi\r\neuismod.\r\nMaecenas\r\neuismod\r\nfermentum\r\ndiam, sit\r\namet\r\ngravida\r\nmagna\r\nsuscipit a.\r\nQuisque\r\nconsectetur\r\narcu eu\r\nnunc\r\nsodales\r\nscelerisque.\r\nNulla non\r\ntincidunt\r\nnulla.\r\nPellentesqu\r\ne ut tortor\r\nvel enim\r\nconvallis\r\nmalesuada.\r\nAliquam\r\nultricies\r\nbibendum\r\nultrices.\r\nMauris\r\nrutrum ac\r\nnisl vel\r\nluctus.\r\nDonec quis\r\nnibh vitae\r\norci ultricies\r\ngravida.\r\nAliquam\r\nvitae velit\r\nporttitor\r\nlorem\r\nbibendum\r\nfringilla\r\nvolutpat a\r\neros.\r\nCurabitur at\r\ncommodo\r\ntortor. Etiam\r\nultricies,\r\nneque et\r\niaculis\r\neuismod,\r\ndiam ligula\r\nluctus mi,\r\nvitae\r\nlobortis felis\r\nlorem eu\r\nnulla. Sed a\r\nsemper ex.\r\nInterdum et\r\nmalesuada\r\nfames ac\r\nante ipsum\r\nprimis in\r\nfaucibus.\r\nNulla\r\nmauris elit,\r\npulvinar ac\r\ntortor et,\r\nluctus\r\nhendrerit\r\nnisl. In\r\negestas\r\nauctor urna\r\nvitae\r\nlaoreet.\r\nPraesent\r\nbibendum\r\negestas\r\nconvallis.\r\nProin non\r\nsuscipit\r\ntellus.\r\nNullam at\r\nnibh in urna\r\nlaoreet\r\nsodales non\r\nvel tellus.\r\nDonec in\r\nenim dui.\r\nPhasellus\r\nquis quam\r\ntincidunt,\r\npellentesqu\r\ne lorem ac,\r\nscelerisque\r\nneque.\r\nInteger nec\r\ntempus\r\nurna. Donec\r\nelit massa,\r\neleifend eu\r\nsapien sit\r\namet,\r\nmollis\r\npellentesqu\r\ne est.\r\nNullam\r\ntristique\r\ntellus\r\niaculis arcu\r\nconsectetur\r\npretium.\r\nSed\r\nvenenatis\r\nconvallis\r\nscelerisque.\r\nSuspendiss\r\ne varius\r\nurna sit\r\namet purus\r\naccumsan,\r\nid ultricies\r\nerat\r\nefficitur.\r\nCras non\r\nipsum eget\r\nnulla\r\nefficitur\r\ncommodo\r\nsit amet non\r\nlacus. Proin\r\nviverra enim\r\nsit amet\r\nenim\r\ntempus\r\nullamcorper\r\n. Class\r\naptent taciti\r\nsociosqu ad\r\nlitora\r\ntorquent per\r\nconubia\r\nnostra, per\r\ninceptos\r\nhimenaeos.\r\nDuis ac\r\nmassa\r\ninterdum,\r\ngravida ex\r\negestas,\r\nfinibus\r\npurus. Nunc\r\nconsectetur\r\ncommodo\r\nlacus, ac\r\nconvallis\r\nquam\r\nlobortis eu.\r\nSed\r\nconvallis\r\ntempor\r\ncommodo.\r\nNulla sed\r\nconvallis\r\nmauris.\r\nDonec\r\nvenenatis\r\nnisi est, ac\r\nullamcorper\r\nmi pretium\r\nquis. Donec\r\nvitae eros at\r\nipsum\r\ninterdum\r\nscelerisque\r\nnec vitae\r\nnisi. Sed\r\nvestibulum\r\nerat ac\r\nbibendum\r\ndapibus.\r\nMorbi nec\r\nelit id quam\r\ntristique\r\ncursus id\r\nsed sem.\r\nPraesent\r\nnon ante\r\nenim.\r\nPellentesqu\r\ne habitant\r\nmorbi\r\ntristique\r\nsenectus et\r\nnetus et\r\nmalesuada\r\nfames ac\r\nturpis\r\negestas.\r\nPraesent\r\nnon mauris\r\ndui.\r\nAliquam\r\nrhoncus\r\nmattis ante\r\nsed\r\nvenenatis.\r\nVivamus\r\nvehicula\r\nsed sapien\r\nsed dictum.\r\nIn aliquet,\r\nurna\r\nefficitur\r\ntincidunt\r\nlobortis,\r\nnibh justo\r\ntristique\r\npurus, sed\r\nvolutpat\r\nrisus magna\r\net\r\nlibero.Susp\r\nendisse\r\nlectus justo,\r\nvarius eget\r\narcu et,\r\nsemper\r\nlaoreet erat.\r\nQuisque\r\neget lacus\r\nornare,\r\npellentesqu\r\ne erat sit\r\namet,\r\nvulputate\r\nfelis. Duis\r\nluctus,\r\nmassa a\r\npellentesqu\r\ne mollis,\r\nmassa elit\r\nconvallis\r\nmi, vel\r\nbibendum\r\nex ex eu\r\npurus.\r\nSuspendiss\r\ne vel\r\nfermentum\r\nurna, ac\r\ncommodo\r\nenim.\r\nMauris\r\ntincidunt\r\ncursus elit,\r\na volutpat\r\nlibero\r\ncommodo\r\net. Etiam\r\ndapibus\r\nlibero\r\nvenenatis\r\ntellus\r\nlobortis, vel\r\nlacinia elit\r\nfaucibus.\r\nMaecenas\r\nsemper sed\r\nquam quis\r\nfinibus.\r\nInteger\r\nefficitur,\r\nlibero\r\nimperdiet\r\nsollicitudin\r\ncommodo,\r\nelit arcu\r\nvulputate\r\nest, eget\r\nfinibus mi\r\nurna sit\r\namet\r\nmagna.\r\nCras\r\nullamcorper\r\nconsequat\r\nornare.\r\nFusce\r\nconvallis\r\nnunc vel\r\nrisus\r\ncursus, at\r\nmaximus\r\nligula\r\ncursus.\r\nPellentesqu\r\ne vulputate\r\nrisus libero,\r\neget cursus\r\nnibh\r\nsodales\r\nsed. Donec\r\naccumsan\r\nsem et\r\nmassa\r\nsemper, id\r\ndignissim\r\nvelit\r\nvehicula.Cr\r\nas cursus\r\nipsum ac\r\nerat\r\nvehicula,\r\nnec iaculis\r\npurus\r\ndictum.\r\nQuisque\r\nlacinia elit\r\nvitae leo\r\ndictum, vel\r\ndignissim\r\nvelit\r\ndapibus.\r\nAenean sem\r\nnisi,\r\nfaucibus\r\ninterdum\r\njusto eu,\r\neuismod\r\nporttitor ex.\r\nMorbi et\r\nlectus\r\nlectus. Duis\r\nneque felis,\r\nsuscipit at\r\nscelerisque\r\neu,\r\nscelerisque\r\nid orci.\r\nCurabitur et\r\nplacerat\r\nipsum.\r\nProin\r\ngravida\r\nsapien nisl,\r\net varius\r\nipsum\r\nmollis nec.\r\nQuisque\r\ndignissim\r\nconsectetur\r\nfeugiat.\r\nAenean\r\neros purus,\r\nlaoreet\r\ninterdum\r\nrutrum at,\r\naliquet sit\r\namet\r\nlectus.\r\nDonec\r\ngravida\r\nlorem ut\r\ntincidunt\r\nlaoreet.\r\nDonec\r\nconsequat\r\nviverra\r\nligula, in\r\naccumsan\r\nmi\r\nbibendum\r\nscelerisque.\r\nQuisque ac\r\nrisus justo.\r\nMorbi\r\nmagna\r\narcu,\r\negestas nec\r\nluctus\r\ncommodo,\r\ncursus eget\r\nnunc.\r\nVivamus\r\neuismod\r\nlorem ex, et\r\nmaximus\r\nfelis\r\nhendrerit\r\neget.\r\nNullam\r\nullamcorper\r\neuismod\r\nligula, et\r\niaculis\r\nligula\r\nultricies a.\r\nFusce\r\naliquam,\r\nenim vel\r\nfermentum\r\nultrices, elit\r\nquam\r\nsemper\r\nerat, vitae\r\nsemper velit\r\naugue non\r\nmagna.Quis\r\nque\r\nmaximus\r\nsemper\r\narcu, id\r\npellentesqu\r\ne est\r\ntempus a.\r\nPhasellus\r\nlacus elit,\r\nauctor sit\r\namet lacinia\r\na, dapibus\r\nvitae velit.\r\nPhasellus ut\r\npharetra\r\njusto, ut\r\nultricies\r\nerat. Sed\r\nmolestie\r\nsapien vel\r\ninterdum\r\nlobortis.\r\nNulla\r\nfacilisi.\r\nVestibulum\r\nante ipsum\r\nprimis in\r\nfaucibus\r\norci luctus\r\net ultrices\r\nposuere\r\ncubilia\r\ncurae; Nulla\r\nnec mauris\r\nquis nisi\r\nvulputate\r\ngravida quis\r\nnec\r\nvelit.Nam et\r\ncongue\r\nipsum.\r\nNulla vel elit\r\nnon dolor\r\nmollis\r\naliquet vel\r\nat magna.\r\nPellentesqu\r\ne nec\r\nfacilisis elit.\r\nIn vulputate\r\nquis sem\r\nporta\r\nsuscipit.\r\nNullam sed\r\nex ornare\r\nnibh\r\nsuscipit\r\nmattis quis\r\nnon lacus.\r\nMauris vel\r\nex urna.\r\nVivamus\r\nultricies\r\nsapien sit\r\namet sapien\r\nvehicula\r\ngravida.\r\nDonec\r\nfeugiat\r\nvolutpat\r\nquam.\r\nVestibulum\r\nauctor\r\ndictum nisl,\r\nid hendrerit\r\nmetus\r\nullamcorper\r\nsed. Nulla\r\nmaximus\r\nlacus vel\r\nmollis\r\nmaximus.\r\nNulla\r\nlaoreet\r\nplacerat\r\nquam eu\r\nviverra.\r\nEtiam\r\nfeugiat\r\naccumsan\r\nnisl a\r\ncondimentu\r\nm. Sed\r\nultricies\r\nante ante,\r\nac auctor\r\nligula\r\ngravida nec.\r\nPraesent a\r\nneque\r\ndignissim,\r\nsagittis felis\r\nsit amet,\r\ncondimentu\r\nm turpis.\r\nFusce at leo\r\nvel est\r\nblandit\r\nmalesuada.\r\nPellentesqu\r\ne et neque\r\nnon metus\r\npellentesqu\r\ne imperdiet.\r\nPraesent\r\npellentesqu\r\ne lacinia\r\nlorem, et\r\ntristique\r\ntellus\r\nefficitur id.\r\nSuspendiss\r\ne aliquet\r\nultricies\r\njusto vitae\r\ninterdum.\r\nCras\r\ntristique\r\nviverra\r\nquam, eget\r\ngravida mi\r\nfermentum\r\nimperdiet.\r\nSed\r\nimperdiet\r\nvitae purus\r\nut volutpat.\r\nNulla\r\nlacinia elit\r\nin\r\nfermentum\r\nconsectetur\r\n. Phasellus\r\ncommodo\r\nut nisl sit\r\namet\r\nsagittis.\r\nDuis ac\r\nornare orci.\r\nVivamus vel\r\nenim\r\nposuere,\r\npharetra ex\r\nvel,\r\nelementum\r\nest.\r\nVestibulum\r\ncommodo\r\nluctus\r\nmetus eget\r\nmaximus.\r\nSuspendiss\r\ne a nulla a\r\nodio\r\neleifend\r\nfaucibus.\r\nSuspendiss\r\ne semper\r\nlacus non\r\nporttitor\r\naliquet.\r\nCras ac\r\nscelerisque\r\nmagna, et\r\npulvinar\r\njusto.\r\nInteger\r\ncursus\r\npulvinar\r\nfringilla.\r\nMauris\r\nimperdiet\r\nnibh sit\r\namet\r\ntempor\r\nlaoreet.\r\nMorbi\r\ntincidunt\r\ntortor ex, sit\r\namet\r\nmaximus\r\npurus\r\ntristique\r\nquis.\r\nQuisque\r\nsed\r\nhendrerit\r\nvelit. Mauris\r\nmattis nibh\r\nut eros\r\nluctus, eget\r\nmattis\r\nmassa\r\nauctor.\r\nPhasellus\r\neu neque at\r\naugue\r\ngravida\r\nsagittis nec\r\nnon tortor.\r\nEtiam\r\nporttitor\r\nsem\r\nsodales mi\r\nullamcorper\r\ngravida. In\r\nin dictum\r\norci. In vitae\r\nvestibulum\r\nquam. Cras\r\naugue eros,\r\ntincidunt ac\r\nelit posuere,\r\nsollicitudin\r\nefficitur\r\nlectus.\r\nPraesent\r\nquis\r\nsodales\r\nnisl. Proin\r\nsit amet\r\nmolestie\r\nest. In\r\ncommodo\r\nmauris vel\r\nmauris\r\nefficitur,\r\nnec mollis\r\nmauris\r\nsagittis.\r\nCras ligula\r\nnibh,\r\negestas sit\r\namet eros\r\nin, lacinia\r\ntristique\r\nmagna.\r\nCras risus\r\nlibero,\r\nlacinia eget\r\nlibero vitae,\r\nmaximus\r\naliquet\r\nnibh. Mauris\r\nid sodales\r\npurus, vitae\r\ndictum\r\nlectus. Cras\r\nconsectetur\r\nligula velit,\r\ntempus\r\npulvinar\r\nlacus\r\nporttitor\r\nvitae.\r\nPhasellus\r\neget tellus\r\nipsum.\r\nDonec\r\ninterdum\r\nlaoreet elit\r\nnon\r\nvestibulum.\r\nCras sed\r\nurna\r\nullamcorper\r\n, aliquam\r\nerat eget,\r\nporta orci.\r\nVestibulum\r\neget congue\r\nnulla. Sed\r\nsem tortor,\r\neuismod at\r\nrutrum id,\r\nsagittis a\r\nnunc. Duis\r\nin nibh\r\nfacilisis,\r\ndignissim\r\npurus ut,\r\nhendrerit\r\nmagna. Sed\r\nsemper\r\nligula id\r\nmassa\r\nelementum,\r\nnon\r\nmalesuada\r\nvelit\r\negestas.\r\nNullam\r\ndictum, mi\r\nnec\r\neuismod\r\nsagittis,\r\nligula leo\r\nullamcorper\r\ndolor, quis\r\nfaucibus\r\nodio metus\r\neget magna.\r\nUt gravida\r\nmetus non\r\nmetus\r\nbibendum\r\nbibendum.\r\nIn sagittis\r\neleifend\r\naliquet.\r\nInterdum et\r\nmalesuada\r\nfames ac\r\nante ipsum\r\nprimis in\r\nfaucibus.\r\nNam mollis\r\nsagittis\r\nfelis, in\r\nfaucibus\r\ntortor\r\npretium vel.\r\nNam nec\r\nenim\r\nmetus.\r\nDonec in\r\naugue arcu.\r\nProin non\r\nlobortis\r\npurus, sit\r\namet lacinia\r\nelit.\r\nSuspendiss\r\ne quis eros\r\ncondimentu\r\nm, blandit\r\njusto sit\r\namet,\r\nlobortis nisl.\r\nSuspendiss\r\ne maximus\r\nmassa sed\r\nurna tempor\r\nornare.\r\nNunc\r\nmalesuada\r\npurus odio,\r\neu luctus\r\nlectus\r\nauctor nec.\r\nMorbi\r\nauctor\r\npellentesqu\r\ne auctor.\r\nSed\r\nullamcorper\r\n, ex vitae\r\naliquam\r\nvulputate,\r\nest diam\r\nfeugiat mi,\r\nid porttitor\r\nlectus orci\r\nac leo.\r\nDonec sit\r\namet velit\r\npulvinar,\r\nvenenatis\r\nturpis ut,\r\ninterdum\r\nligula.\r\nInterdum et\r\nmalesuada\r\nfames ac\r\nante ipsum\r\nprimis in\r\nfaucibus.\r\nVestibulum\r\neu lacus\r\nurna.\r\nMaecenas\r\nsem nulla,\r\naccumsan\r\neu ultricies\r\nsed, tempor\r\nvel magna.\r\nCras aliquet\r\nsollicitudin\r\nsapien ac\r\npulvinar.\r\nPraesent ac\r\nsodales mi.\r\nInteger vitae\r\nmauris\r\nmassa.\r\nMaecenas\r\niaculis orci\r\net faucibus\r\ninterdum.\r\nNunc nec\r\nmaximus\r\nfelis, sed\r\nfinibus\r\nquam.\r\nPellentesqu\r\ne felis\r\nmassa,\r\nvestibulum\r\nin tellus\r\nvitae,\r\ncongue\r\ntincidunt\r\njusto. Nunc\r\nvitae enim\r\nmalesuada,\r\nbibendum\r\nante nec,\r\nvarius\r\ntellus.\r\nPraesent\r\nvitae nisi id\r\nquam\r\nauctor\r\nlacinia at\r\nnon quam.\r\nNam nec\r\nligula sit\r\namet felis\r\nauctor\r\nsagittis.\r\nNunc in\r\nrisus eu\r\nurna varius\r\nlaoreet quis\r\nsit amet\r\nfelis. Morbi\r\nvarius\r\ntempor orci,\r\neu\r\nvestibulum\r\nnunc\r\nvestibulum\r\nac. Nunc\r\nvehicula\r\nvelit\r\neleifend\r\nconsequat\r\nporta.\r\nSuspendiss\r\ne maximus\r\ndapibus\r\norci, in\r\nvulputate\r\nmassa\r\npretium ac.\r\nQuisque\r\nmalesuada\r\naliquet\r\naliquet.";
            var savedStrings = SavedComparisonString.Split("\r\n");

            ITextShaper shaper = new TextShaper(font);
            using var layoutEngine = new TextLayoutEngine(shaper);
            var wrappedLines = layoutEngine.WrapText(
                Lorem20Str,
                11f,
                54,
                ShapingOptions.Full
            );

            List<string> faultyStrings = new();
            List<string> excpectedStrings = new();
            List<int> indiciesOfDifferingString = new();

            for (int i = 0; i < wrappedLines.Count(); i++)
            {
                if (savedStrings[i] != wrappedLines[i])
                {
                    indiciesOfDifferingString.Add(i);
                    faultyStrings.Add(wrappedLines[i]);
                    excpectedStrings.Add(savedStrings[i]);
                }
            }

            if (indiciesOfDifferingString.Count != 0)
            {
                //The start of indicies diverging
                Assert.IsNull(indiciesOfDifferingString[0]);
                Assert.AreEqual(faultyStrings[0], excpectedStrings[0]);
            }

            Assert.AreEqual(0, faultyStrings.Count);
        }
        
        [TestMethod]
        public void WrapText_WithPreExistingWidth_AccountsForIt()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Measure "Hello " to get its width
            var testShaper = new TextShaper(font);
            var shaped = testShaper.Shape("Hello ", ShapingOptions.Default);
            double preWidth = shaped.GetWidthInPoints(11f, testShaper.UnitsPerEm);

            // Act - Add text with pre-existing width, narrow max width
            var lines = layout.WrapText("world test", 11f, preWidth + 50, preWidth);

            // Assert - Should wrap because first line already has content
            Assert.IsTrue(lines.Count >= 1);
        }

        [TestMethod]
        public void WrapText_EmptyString_ReturnsEmptyLine()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapText("", 11f, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual(string.Empty, lines[0]);
        }

        [TestMethod]
        public void WrapText_WithKerning_MeasuresCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act - "AV" has kerning in Roboto
            var withKerning = layout.WrapText("AV", 11f, 1000, ShapingOptions.Default);
            var withoutKerning = layout.WrapText("AV", 11f, 1000, ShapingOptions.None);

            // Assert - Both should be single line, but measured differently
            Assert.AreEqual(1, withKerning.Count);
            Assert.AreEqual(1, withoutKerning.Count);
            Assert.AreEqual("AV", withKerning[0]);
            Assert.AreEqual("AV", withoutKerning[0]);
        }

        #endregion

        #region Rich Text Wrapping Tests

        [TestMethod]
        public void WrapRichText_SingleFragment_BehavesLikeSingleFont()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hello world",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_MultipleFragments_ConcatenatesCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hello ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "world",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 12, Style = MeasurementFontStyles.Bold }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_DifferentFonts_WrapsCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "This is ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "mixed ",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 14, Style = MeasurementFontStyles.Bold }
                },
                new TextFragment
                {
                    Text = "fonts",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act - narrow width to force wrapping
            var lines = layout.WrapRichText(fragments, 80);

            // Assert
            Assert.IsTrue(lines.Count >= 1);

            // Check that we got the expected lines
            Assert.AreEqual("This is mixed", lines[0]);
            Assert.AreEqual("fonts", lines[1]);

            // When joining wrapped lines with spaces, we get back close to original
            string rejoined = string.Join(" ", lines);
            Assert.AreEqual("This is mixed fonts", rejoined);
        }

        [TestMethod]
        public void WrapRichText_WordSpanningFragments_MeasuresCorrectly()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // "Hello" split across two fragments with different fonts
            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Hel",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "lo world",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello world", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_WithLineBreaks_PreservesBreaks()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "Line 1\n",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "Line 2",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(2, lines.Count);
            Assert.AreEqual("Line 1", lines[0]);
            Assert.AreEqual("Line 2", lines[1]);
        }

        [TestMethod]
        public void WrapRichText_EmptyFragments_HandlesGracefully()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "Hello",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                },
                new TextFragment
                {
                    Text = "",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                }
            };

            // Act
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hello", lines[0]);
        }

        [TestMethod]
        public void WrapRichText_NullFragmentList_ReturnsEmptyLine()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            // Act
            var lines = layout.WrapRichText(null, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual(string.Empty, lines[0]);
        }

        #endregion

        #region Font Caching Tests

        [TestMethod]
        public void WrapRichText_SameFontMultipleTimes_UsesCache()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Calibri", FontSubFamily.Regular);
            var shaper = new TextShaper(font);
            var layout = new TextLayoutEngine(shaper);

            var fragments = new List<TextFragment>
            {
                new TextFragment
                {
                    Text = "First ",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 }
                },
                new TextFragment
                {
                    Text = "second ",
                    Font = new MeasurementFont { FontFamily = "Calibri", Size = 11 }
                },
                new TextFragment
                {
                    Text = "third",
                    Font = new MeasurementFont { FontFamily = "Arial", Size = 11 } // Same as first
                }
            };

            // Act - This should use cached shaper for Arial
            var lines = layout.WrapRichText(fragments, 1000);

            // Assert
            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("First second third", lines[0]);
            // Note: We can't easily verify cache usage without exposing internals,
            // but this test documents the expected behavior
        }

        #endregion

        [TestMethod]
        public void WrapText_Continous_Long_Word()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Aptos Narrow", FontSubFamily.Regular);

            var longWord = "pellentesquer";

            ITextShaper shaper = new TextShaper(font);
            using var layoutEngine = new TextLayoutEngine(shaper);
            var wrappedLines = layoutEngine.WrapText(
                longWord,
                11f,
                54,
                ShapingOptions.Full
            );

            Assert.AreEqual("pellentesqu", wrappedLines[0]);
            Assert.AreEqual("er", wrappedLines[1]);
        }
    }
}