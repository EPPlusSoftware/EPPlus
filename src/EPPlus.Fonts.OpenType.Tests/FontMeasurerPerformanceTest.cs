using EPPlus.Fonts.OpenType.FontCache;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;

namespace EPPlus.Fonts.OpenType.Tests
{
    
    /// <summary>
    /// Uses STATesClass to ensure single threaded for performance test
    /// </summary>
    // Terms to Ensure search finds this later: single-thread , single-threaded test , STA , single-threaded apartment
    [TestClass]
    [STATestClass]
    public class FontMeasurerPerformanceTest
    {
        //20 paragraphs of 'lorem ipsum' statistics
        //1706 words
        //11 800 characters (with whitespace)
        //10 095 characters without whitespaces
        const string LoremIpsum20Para = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla pulvinar interdum imperdiet. Praesent ut auctor urna. Phasellus sollicitudin quam vitae est convallis, eu mattis lorem efficitur. Mauris nulla libero, tincidunt id ipsum non, lobortis tristique mauris. Donec ut enim sed enim fermentum molestie vel quis odio. Morbi a fermentum massa, sit amet ultrices est. Aenean ante mi, fermentum nec rhoncus et, vulputate vel sapien. Donec tempus, leo quis luctus rhoncus, augue odio pharetra libero, ac blandit urna turpis sed diam. Vivamus augue purus, eleifend et justo facilisis, imperdiet rhoncus sem. Quisque accumsan pellentesque elit, eget finibus massa accumsan in.\r\n\r\nFusce eu accumsan enim. Cras pulvinar enim vel tellus lacinia, consectetur euismod tortor consectetur. Praesent tincidunt pretium eros, ac auctor magna luctus sed. Ut porta lectus quam, non ornare mauris lacinia sit amet. Nullam egestas dolor quis magna porttitor, ac iaculis nisi hendrerit. Proin at mollis lacus, in porttitor nunc. Aliquam erat volutpat. Sed vel egestas risus, at aliquam arcu. Vestibulum quis lobortis nulla. Etiam pellentesque auctor nulla, eget tincidunt felis rhoncus id. Sed metus ante, efficitur id dui eu, fermentum mollis odio. Phasellus ullamcorper iaculis augue vel consequat. Etiam fringilla euismod interdum. Ut molestie massa id fringilla lobortis. Vestibulum malesuada, ante vel mattis ultrices, sem ante molestie augue, non tristique dui mi non nibh.\r\n\r\nMaecenas dictum, sem eget convallis rhoncus, lacus enim porta neque, in posuere dui ex a sapien. Nam lacus nibh, posuere sed elit eget, condimentum facilisis ligula. Cras consectetur lacus ullamcorper velit aliquet bibendum eget vel nulla. Aenean varius ac erat quis ullamcorper. Donec laoreet arcu a lorem volutpat faucibus. Vivamus vehicula leo ut erat luctus scelerisque. Morbi posuere ex et magna egestas facilisis. Fusce scelerisque volutpat erat bibendum hendrerit. Nam blandit mi ut metus pulvinar, vel tempus lacus euismod. Quisque imperdiet sit amet sapien sed ultricies. Phasellus sodales, ipsum vitae tincidunt facilisis, nulla ligula faucibus felis, eget vehicula ante lacus eu lorem.\r\n\r\nInteger congue diam ac viverra tristique. Curabitur tristique dolor quis quam pretium, et scelerisque quam dictum. Maecenas vitae sodales ligula. Pellentesque maximus diam vel porta convallis. Ut aliquam eros quis porta pellentesque. Fusce in ex ut mi egestas cursus. Aliquam erat volutpat. Cras laoreet condimentum laoreet.\r\n\r\nSed eget facilisis tellus. Morbi viverra odio sed odio placerat mollis. Duis turpis metus, dignissim varius urna quis, viverra dignissim dui. Vivamus viverra at nisi quis convallis. Suspendisse fringilla risus et ante sollicitudin, sed eleifend sem placerat. Proin pretium blandit arcu, eget rhoncus risus hendrerit at. Interdum et malesuada fames ac ante ipsum primis in faucibus. Phasellus vulputate efficitur maximus.\r\n\r\nCras blandit nulla eu nisi auctor tempus. Sed pretium lacus ac magna vestibulum, aliquam faucibus orci luctus. Mauris enim lorem, varius ut ante quis, varius viverra lectus. Fusce blandit nibh vel feugiat efficitur. Donec maximus id justo ac mollis. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla placerat lectus et purus dictum, id congue nisi euismod. Maecenas euismod fermentum diam, sit amet gravida magna suscipit a. Quisque consectetur arcu eu nunc sodales scelerisque. Nulla non tincidunt nulla. Pellentesque ut tortor vel enim convallis malesuada.\r\n\r\nAliquam ultricies bibendum ultrices. Mauris rutrum ac nisl vel luctus. Donec quis nibh vitae orci ultricies gravida. Aliquam vitae velit porttitor lorem bibendum fringilla volutpat a eros. Curabitur at commodo tortor. Etiam ultricies, neque et iaculis euismod, diam ligula luctus mi, vitae lobortis felis lorem eu nulla. Sed a semper ex. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nulla mauris elit, pulvinar ac tortor et, luctus hendrerit nisl. In egestas auctor urna vitae laoreet. Praesent bibendum egestas convallis. Proin non suscipit tellus.\r\n\r\nNullam at nibh in urna laoreet sodales non vel tellus. Donec in enim dui. Phasellus quis quam tincidunt, pellentesque lorem ac, scelerisque neque. Integer nec tempus urna. Donec elit massa, eleifend eu sapien sit amet, mollis pellentesque est. Nullam tristique tellus iaculis arcu consectetur pretium. Sed venenatis convallis scelerisque. Suspendisse varius urna sit amet purus accumsan, id ultricies erat efficitur. Cras non ipsum eget nulla efficitur commodo sit amet non lacus. Proin viverra enim sit amet enim tempus ullamcorper. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Duis ac massa interdum, gravida ex egestas, finibus purus. Nunc consectetur commodo lacus, ac convallis quam lobortis eu. Sed convallis tempor commodo. Nulla sed convallis mauris.\r\n\r\nDonec venenatis nisi est, ac ullamcorper mi pretium quis. Donec vitae eros at ipsum interdum scelerisque nec vitae nisi. Sed vestibulum erat ac bibendum dapibus. Morbi nec elit id quam tristique cursus id sed sem. Praesent non ante enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Praesent non mauris dui. Aliquam rhoncus mattis ante sed venenatis. Vivamus vehicula sed sapien sed dictum. In aliquet, urna efficitur tincidunt lobortis, nibh justo tristique purus, sed volutpat risus magna et libero.\r\n\r\nSuspendisse lectus justo, varius eget arcu et, semper laoreet erat. Quisque eget lacus ornare, pellentesque erat sit amet, vulputate felis. Duis luctus, massa a pellentesque mollis, massa elit convallis mi, vel bibendum ex ex eu purus. Suspendisse vel fermentum urna, ac commodo enim. Mauris tincidunt cursus elit, a volutpat libero commodo et. Etiam dapibus libero venenatis tellus lobortis, vel lacinia elit faucibus. Maecenas semper sed quam quis finibus. Integer efficitur, libero imperdiet sollicitudin commodo, elit arcu vulputate est, eget finibus mi urna sit amet magna. Cras ullamcorper consequat ornare. Fusce convallis nunc vel risus cursus, at maximus ligula cursus. Pellentesque vulputate risus libero, eget cursus nibh sodales sed. Donec accumsan sem et massa semper, id dignissim velit vehicula.\r\n\r\nCras cursus ipsum ac erat vehicula, nec iaculis purus dictum. Quisque lacinia elit vitae leo dictum, vel dignissim velit dapibus. Aenean sem nisi, faucibus interdum justo eu, euismod porttitor ex. Morbi et lectus lectus. Duis neque felis, suscipit at scelerisque eu, scelerisque id orci. Curabitur et placerat ipsum. Proin gravida sapien nisl, et varius ipsum mollis nec. Quisque dignissim consectetur feugiat. Aenean eros purus, laoreet interdum rutrum at, aliquet sit amet lectus. Donec gravida lorem ut tincidunt laoreet. Donec consequat viverra ligula, in accumsan mi bibendum scelerisque. Quisque ac risus justo. Morbi magna arcu, egestas nec luctus commodo, cursus eget nunc. Vivamus euismod lorem ex, et maximus felis hendrerit eget. Nullam ullamcorper euismod ligula, et iaculis ligula ultricies a. Fusce aliquam, enim vel fermentum ultrices, elit quam semper erat, vitae semper velit augue non magna.\r\n\r\nQuisque maximus semper arcu, id pellentesque est tempus a. Phasellus lacus elit, auctor sit amet lacinia a, dapibus vitae velit. Phasellus ut pharetra justo, ut ultricies erat. Sed molestie sapien vel interdum lobortis. Nulla facilisi. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Nulla nec mauris quis nisi vulputate gravida quis nec velit.\r\n\r\nNam et congue ipsum. Nulla vel elit non dolor mollis aliquet vel at magna. Pellentesque nec facilisis elit. In vulputate quis sem porta suscipit. Nullam sed ex ornare nibh suscipit mattis quis non lacus. Mauris vel ex urna. Vivamus ultricies sapien sit amet sapien vehicula gravida. Donec feugiat volutpat quam. Vestibulum auctor dictum nisl, id hendrerit metus ullamcorper sed. Nulla maximus lacus vel mollis maximus. Nulla laoreet placerat quam eu viverra. Etiam feugiat accumsan nisl a condimentum. Sed ultricies ante ante, ac auctor ligula gravida nec. Praesent a neque dignissim, sagittis felis sit amet, condimentum turpis.\r\n\r\nFusce at leo vel est blandit malesuada. Pellentesque et neque non metus pellentesque imperdiet. Praesent pellentesque lacinia lorem, et tristique tellus efficitur id. Suspendisse aliquet ultricies justo vitae interdum. Cras tristique viverra quam, eget gravida mi fermentum imperdiet. Sed imperdiet vitae purus ut volutpat. Nulla lacinia elit in fermentum consectetur. Phasellus commodo ut nisl sit amet sagittis. Duis ac ornare orci.\r\n\r\nVivamus vel enim posuere, pharetra ex vel, elementum est. Vestibulum commodo luctus metus eget maximus. Suspendisse a nulla a odio eleifend faucibus. Suspendisse semper lacus non porttitor aliquet. Cras ac scelerisque magna, et pulvinar justo. Integer cursus pulvinar fringilla. Mauris imperdiet nibh sit amet tempor laoreet. Morbi tincidunt tortor ex, sit amet maximus purus tristique quis. Quisque sed hendrerit velit. Mauris mattis nibh ut eros luctus, eget mattis massa auctor. Phasellus eu neque at augue gravida sagittis nec non tortor. Etiam porttitor sem sodales mi ullamcorper gravida.\r\n\r\nIn in dictum orci. In vitae vestibulum quam. Cras augue eros, tincidunt ac elit posuere, sollicitudin efficitur lectus. Praesent quis sodales nisl. Proin sit amet molestie est. In commodo mauris vel mauris efficitur, nec mollis mauris sagittis. Cras ligula nibh, egestas sit amet eros in, lacinia tristique magna. Cras risus libero, lacinia eget libero vitae, maximus aliquet nibh. Mauris id sodales purus, vitae dictum lectus. Cras consectetur ligula velit, tempus pulvinar lacus porttitor vitae. Phasellus eget tellus ipsum.\r\n\r\nDonec interdum laoreet elit non vestibulum. Cras sed urna ullamcorper, aliquam erat eget, porta orci. Vestibulum eget congue nulla. Sed sem tortor, euismod at rutrum id, sagittis a nunc. Duis in nibh facilisis, dignissim purus ut, hendrerit magna. Sed semper ligula id massa elementum, non malesuada velit egestas. Nullam dictum, mi nec euismod sagittis, ligula leo ullamcorper dolor, quis faucibus odio metus eget magna. Ut gravida metus non metus bibendum bibendum. In sagittis eleifend aliquet.\r\n\r\nInterdum et malesuada fames ac ante ipsum primis in faucibus. Nam mollis sagittis felis, in faucibus tortor pretium vel. Nam nec enim metus. Donec in augue arcu. Proin non lobortis purus, sit amet lacinia elit. Suspendisse quis eros condimentum, blandit justo sit amet, lobortis nisl. Suspendisse maximus massa sed urna tempor ornare. Nunc malesuada purus odio, eu luctus lectus auctor nec. Morbi auctor pellentesque auctor. Sed ullamcorper, ex vitae aliquam vulputate, est diam feugiat mi, id porttitor lectus orci ac leo.\r\n\r\nDonec sit amet velit pulvinar, venenatis turpis ut, interdum ligula. Interdum et malesuada fames ac ante ipsum primis in faucibus. Vestibulum eu lacus urna. Maecenas sem nulla, accumsan eu ultricies sed, tempor vel magna. Cras aliquet sollicitudin sapien ac pulvinar. Praesent ac sodales mi. Integer vitae mauris massa. Maecenas iaculis orci et faucibus interdum.\r\n\r\nNunc nec maximus felis, sed finibus quam. Pellentesque felis massa, vestibulum in tellus vitae, congue tincidunt justo. Nunc vitae enim malesuada, bibendum ante nec, varius tellus. Praesent vitae nisi id quam auctor lacinia at non quam. Nam nec ligula sit amet felis auctor sagittis. Nunc in risus eu urna varius laoreet quis sit amet felis. Morbi varius tempor orci, eu vestibulum nunc vestibulum ac. Nunc vehicula velit eleifend consequat porta. Suspendisse maximus dapibus orci, in vulputate massa pretium ac. Quisque malesuada aliquet aliquet.";
        
        /// <summary>
        /// String from file before optimization to ensure optimization did not change wrapping.
        /// </summary>
        const string UnOptimizedOriginalWrappingString = "Lorem\r\nipsum\r\ndolor sit\r\namet,\r\nconsectetur\r\nadipiscin\r\ng elit.\r\nNulla\r\npulvinar\r\ninterdum\r\nimperdiet.\r\nPraesent\r\nut auctor\r\nurna.\r\nPhasellus\r\nsollicitudi\r\nn quam\r\nvitae est\r\nconvallis,\r\neu\r\nmattis lorem\r\nefficitur.\r\nMauris\r\nnulla\r\nlibero,\r\ntincidunt id\r\nipsum\r\nnon, lobortis\r\ntristique\r\nmauris.\r\nDonec\r\nut enim\r\nsed enim\r\nferment\r\num\r\nmolestie vel\r\nquis odio.\r\nMorbi a\r\nfermentu\r\nm massa,\r\nsit amet\r\nultrices\r\nest.\r\nAenean ante\r\nmi,\r\nfermentum\r\nnec\r\nrhoncus et,\r\nvulputate\r\nvel\r\nsapien.\r\nDonec\r\ntempus, leo\r\nquis luctus\r\nrhoncus,\r\naugue\r\nodio\r\npharetra\r\nlibero, ac\r\nblandit urna\r\nturpis\r\nsed diam.\r\nVivamus\r\naugue\r\npurus,\r\neleifend et\r\njusto\r\nfacilisis,\r\nimperdiet\r\nrhoncus\r\nsem.\r\nQuisque\r\naccumsan\r\npellente\r\nsque elit,\r\neget\r\nfinibus\r\nmassa\r\naccumsan in.\r\n\r\nFusce\r\neu\r\naccumsan\r\nenim. Cras\r\npulvinar\r\nenim vel\r\ntellus\r\nlacinia,\r\nconsectetu\r\nr\r\neuismod tortor\r\nconsect\r\netur.\r\nPraesent\r\ntincidunt\r\npretium\r\neros, ac\r\nauctor\r\nmagna luctus\r\nsed. Ut\r\nporta\r\nlectus\r\nquam, non\r\nornare\r\nmauris\r\nlacinia sit\r\namet.\r\nNullam\r\negestas dolor\r\nquis\r\nmagna\r\nporttitor, ac\r\niaculis\r\nnisi\r\nhendrerit.\r\nProin at\r\nmollis lacus,\r\nin\r\nporttitor nunc.\r\nAliquam\r\nerat\r\nvolutpat.\r\nSed vel\r\negestas\r\nrisus, at\r\naliquam\r\narcu.\r\nVestibulum\r\nquis\r\nlobortis nulla.\r\nEtiam\r\npellentesq\r\nue auctor\r\nnulla,\r\neget\r\ntincidunt felis\r\nrhoncus\r\nid. Sed\r\nmetus\r\nante,\r\nefficitur id\r\ndui eu,\r\nfermentum\r\nmollis\r\nodio.\r\nPhasellus\r\nullamcorp\r\ner iaculis\r\naugue\r\nvel\r\nconsequat.\r\nEtiam\r\nfringilla\r\neuismod\r\ninterdum. Ut\r\nmolestie\r\nmassa\r\nid\r\nfringilla\r\nlobortis.\r\nVestibulum\r\nmalesuada,\r\nante vel\r\nmattis\r\nultrices,\r\nsem ante\r\nmolestie\r\naugue,\r\nnon\r\ntristique dui\r\nmi non\r\nnibh.\r\n\r\nMaecen\r\nas\r\ndictum, sem\r\neget\r\nconvallis\r\nrhoncus,\r\nlacus enim\r\nporta\r\nneque, in\r\nposuere\r\ndui ex a\r\nsapien.\r\nNam\r\nlacus nibh,\r\nposuere\r\nsed elit\r\neget,\r\ncondimentu\r\nm facilisis\r\nligula.\r\nCras\r\nconsectetur\r\nlacus\r\nullamcorp\r\ner velit\r\naliquet\r\nbibendum\r\neget vel\r\nnulla.\r\nAenean\r\nvarius ac\r\nerat quis\r\nullamcor\r\nper.\r\nDonec laoreet\r\narcu a\r\nlorem\r\nvolutpat\r\nfaucibus.\r\nVivamus\r\nvehicula\r\nleo ut erat\r\nluctus\r\nscelerisq\r\nue. Morbi\r\nposuere\r\nex et\r\nmagna\r\negestas\r\nfacilisis.\r\nFusce\r\nscelerisque\r\nvolutpat\r\nerat\r\nbibendum\r\nhendrerit.\r\nNam\r\nblandit mi ut\r\nmetus\r\npulvinar, vel\r\ntempus\r\nlacus\r\neuismod.\r\nQuisque\r\nimperdie\r\nt sit amet\r\nsapien\r\nsed\r\nultricies.\r\nPhasellus\r\nsodales,\r\nipsum\r\nvitae\r\ntincidunt\r\nfacilisis, nulla\r\nligula\r\nfaucibus\r\nfelis, eget\r\nvehicula\r\nante\r\nlacus eu\r\nlorem.\r\n\r\nInteger\r\ncongue\r\ndiam ac\r\nviverra\r\ntristique.\r\nCurabitur\r\ntristique\r\ndolor\r\nquis quam\r\npretium,\r\net\r\nscelerisque\r\nquam\r\ndictum.\r\nMaecenas\r\nvitae\r\nsodales ligula.\r\nPellente\r\nsque\r\nmaximus\r\ndiam vel\r\nporta\r\nconvallis. Ut\r\naliquam\r\neros\r\nquis porta\r\npellentes\r\nque.\r\nFusce in ex ut\r\nmi\r\negestas\r\ncursus.\r\nAliquam erat\r\nvolutpat.\r\nCras\r\nlaoreet\r\ncondimentu\r\nm laoreet.\r\n\r\nSed eget\r\nfacilisis\r\ntellus.\r\nMorbi\r\nviverra odio\r\nsed odio\r\nplacerat\r\nmollis.\r\nDuis\r\nturpis metus,\r\ndignissi\r\nm varius\r\nurna quis,\r\nviverra\r\ndignissim\r\ndui.\r\nVivamus\r\nviverra at\r\nnisi quis\r\nconvallis.\r\nSuspendi\r\nsse\r\nfringilla risus\r\net ante\r\nsollicitudin\r\n, sed\r\neleifend sem\r\nplacerat.\r\nProin\r\npretium\r\nblandit\r\narcu, eget\r\nrhoncus\r\nrisus\r\nhendrerit at.\r\nInterdu\r\nm et\r\nmalesuada\r\nfames ac\r\nante\r\nipsum primis\r\nin\r\nfaucibus.\r\nPhasellus\r\nvulputate\r\nefficitur\r\nmaximus.\r\n\r\nCras\r\nblandit\r\nnulla eu nisi\r\nauctor\r\ntempus.\r\nSed\r\npretium\r\nlacus ac\r\nmagna\r\nvestibulum,\r\naliquam\r\nfaucibus\r\norci\r\nluctus.\r\nMauris enim\r\nlorem,\r\nvarius ut\r\nante quis,\r\nvarius\r\nviverra\r\nlectus.\r\nFusce\r\nblandit nibh\r\nvel feugiat\r\nefficitur.\r\nDonec\r\nmaximu\r\ns id justo\r\nac\r\nmollis.\r\nVestibulum\r\nante ipsum\r\nprimis in\r\nfaucibus\r\norci\r\nluctus et\r\nultrices\r\nposuere\r\ncubilia\r\ncurae; Nulla\r\nplacerat\r\nlectus et\r\npurus\r\ndictum, id\r\ncongue\r\nnisi\r\neuismod.\r\nMaecenas\r\neuismod\r\nferment\r\num diam,\r\nsit amet\r\ngravida\r\nmagna\r\nsuscipit\r\na.\r\nQuisque\r\nconsectetur\r\narcu eu\r\nnunc\r\nsodales\r\nscelerisque.\r\nNulla\r\nnon\r\ntincidunt nulla.\r\nPellente\r\nsque ut\r\ntortor vel\r\nenim\r\nconvallis\r\nmalesuada.\r\n\r\nAliquam\r\nultricies\r\nbibendu\r\nm ultrices.\r\nMauris\r\nrutrum\r\nac nisl vel\r\nluctus.\r\nDonec\r\nquis nibh\r\nvitae orci\r\nultricies\r\ngravida.\r\nAliquam\r\nvitae velit\r\nporttitor\r\nlorem\r\nbibendum\r\nfringilla\r\nvolutpat\r\na eros.\r\nCurabitur\r\nat\r\ncommodo\r\ntortor. Etiam\r\nultricies,\r\nneque et\r\niaculis\r\neuismod,\r\ndiam\r\nligula\r\nluctus mi,\r\nvitae\r\nlobortis felis\r\nlorem eu\r\nnulla. Sed\r\na\r\nsemper ex.\r\nInterdum et\r\nmalesua\r\nda fames\r\nac ante\r\nipsum\r\nprimis in\r\nfaucibus.\r\nNulla\r\nmauris\r\nelit,\r\npulvinar ac\r\ntortor et,\r\nluctus\r\nhendrerit nisl.\r\nIn\r\negestas\r\nauctor urna\r\nvitae\r\nlaoreet.\r\nPraesent\r\nbibendum\r\negestas\r\nconvallis.\r\nProin non\r\nsuscipit\r\ntellus.\r\n\r\nNullam\r\nat nibh\r\nin urna\r\nlaoreet\r\nsodales\r\nnon vel\r\ntellus.\r\nDonec in\r\nenim dui.\r\nPhasellus\r\nquis\r\nquam\r\ntincidunt,\r\npellentesqu\r\ne lorem\r\nac,\r\nscelerisque\r\nneque.\r\nInteger nec\r\ntempus\r\nurna.\r\nDonec elit\r\nmassa,\r\neleifend eu\r\nsapien\r\nsit amet,\r\nmollis\r\npellentes\r\nque est.\r\nNullam\r\ntristique\r\ntellus\r\niaculis arcu\r\nconsectet\r\nur\r\npretium. Sed\r\nvenenatis\r\nconvallis\r\nsceleris\r\nque.\r\nSuspendisse\r\nvarius\r\nurna sit\r\namet\r\npurus\r\naccumsan, id\r\nultricies\r\nerat\r\nefficitur.\r\nCras non\r\nipsum eget\r\nnulla\r\nefficitur\r\ncommodo\r\nsit amet\r\nnon\r\nlacus. Proin\r\nviverra\r\nenim sit\r\namet\r\nenim\r\ntempus\r\nullamcorper.\r\nClass\r\naptent\r\ntaciti\r\nsociosqu ad\r\nlitora\r\ntorquent per\r\nconubia\r\nnostra,\r\nper\r\ninceptos\r\nhimenaeos.\r\nDuis ac\r\nmassa\r\ninterdum,\r\ngravida ex\r\negestas,\r\nfinibus\r\npurus.\r\nNunc\r\nconsectetur\r\ncommod\r\no lacus,\r\nac\r\nconvallis quam\r\nlobortis\r\neu. Sed\r\nconvallis\r\ntempor\r\ncommo\r\ndo. Nulla\r\nsed\r\nconvallis\r\nmauris.\r\n\r\nDonec\r\nvenenatis\r\nnisi est,\r\nac\r\nullamcorper\r\nmi\r\npretium quis.\r\nDonec\r\nvitae eros\r\nat ipsum\r\ninterdu\r\nm\r\nscelerisque nec\r\nvitae\r\nnisi. Sed\r\nvestibulum\r\nerat ac\r\nbibendum\r\ndapibus.\r\nMorbi\r\nnec elit id\r\nquam\r\ntristique\r\ncursus id\r\nsed sem.\r\nPraesent\r\nnon ante\r\nenim.\r\nPellentesq\r\nue\r\nhabitant morbi\r\ntristique\r\nsenectu\r\ns et netus\r\net\r\nmalesuada\r\nfames ac\r\nturpis\r\negestas.\r\nPraesent\r\nnon\r\nmauris dui.\r\nAliquam\r\nrhoncus\r\nmattis\r\nante sed\r\nvenenatis.\r\nVivamus\r\nvehicula\r\nsed\r\nsapien sed\r\ndictum. In\r\naliquet,\r\nurna\r\nefficitur\r\ntincidunt\r\nlobortis,\r\nnibh justo\r\ntristique\r\npurus,\r\nsed\r\nvolutpat risus\r\nmagna\r\net libero.\r\n\r\nSuspend\r\nisse\r\nlectus justo,\r\nvarius\r\neget arcu\r\net,\r\nsemper laoreet\r\nerat.\r\nQuisque\r\neget lacus\r\nornare,\r\npellentes\r\nque erat\r\nsit amet,\r\nvulputat\r\ne felis.\r\nDuis\r\nluctus, massa\r\na\r\npellentesque\r\nmollis,\r\nmassa elit\r\nconvallis\r\nmi, vel\r\nbibendum\r\nex ex eu\r\npurus.\r\nSuspendi\r\nsse vel\r\nfermentum\r\nurna, ac\r\ncommo\r\ndo enim.\r\nMauris\r\ntincidunt\r\ncursus\r\nelit, a\r\nvolutpat\r\nlibero\r\ncommodo et.\r\nEtiam\r\ndapibus\r\nlibero\r\nvenenatis\r\ntellus\r\nlobortis, vel\r\nlacinia elit\r\nfaucibus\r\n.\r\nMaecenas\r\nsemper sed\r\nquam\r\nquis\r\nfinibus. Integer\r\nefficitur,\r\nlibero\r\nimperdiet\r\nsollicitudi\r\nn\r\ncommodo, elit\r\narcu\r\nvulputate\r\nest, eget\r\nfinibus mi\r\nurna sit\r\namet\r\nmagna.\r\nCras\r\nullamcorper\r\nconsequat\r\nornare.\r\nFusce\r\nconvallis\r\nnunc vel\r\nrisus\r\ncursus, at\r\nmaximus\r\nligula\r\ncursus.\r\nPellentesqu\r\ne\r\nvulputate risus\r\nlibero,\r\neget\r\ncursus nibh\r\nsodales\r\nsed.\r\nDonec\r\naccumsan sem\r\net\r\nmassa\r\nsemper, id\r\ndignissim\r\nvelit\r\nvehicula.\r\n\r\nCras\r\ncursus\r\nipsum ac\r\nerat\r\nvehicula, nec\r\niaculis\r\npurus\r\ndictum.\r\nQuisque\r\nlacinia elit\r\nvitae leo\r\ndictum,\r\nvel\r\ndignissim velit\r\ndapibus.\r\nAenean\r\nsem\r\nnisi,\r\nfaucibus\r\ninterdum justo\r\neu,\r\neuismod\r\nporttitor ex.\r\nMorbi et\r\nlectus\r\nlectus.\r\nDuis neque\r\nfelis,\r\nsuscipit at\r\nscelerisq\r\nue eu,\r\nscelerisque\r\nid orci.\r\nCurabitur\r\net\r\nplacerat\r\nipsum. Proin\r\ngravida\r\nsapien\r\nnisl, et\r\nvarius ipsum\r\nmollis\r\nnec.\r\nQuisque\r\ndignissim\r\nconsectetu\r\nr feugiat.\r\nAenean\r\neros\r\npurus,\r\nlaoreet\r\ninterdum\r\nrutrum at,\r\naliquet sit\r\namet\r\nlectus.\r\nDonec\r\ngravida lorem\r\nut\r\ntincidunt\r\nlaoreet.\r\nDonec\r\nconsequat\r\nviverra\r\nligula, in\r\naccumsan mi\r\nbibendu\r\nm\r\nscelerisque.\r\nQuisque ac\r\nrisus\r\njusto. Morbi\r\nmagna\r\narcu,\r\negestas nec\r\nluctus\r\ncommod\r\no, cursus\r\neget\r\nnunc.\r\nVivamus\r\neuismod\r\nlorem ex, et\r\nmaximu\r\ns felis\r\nhendrerit\r\neget.\r\nNullam\r\nullamcorper\r\neuismod\r\nligula, et\r\niaculis\r\nligula\r\nultricies a.\r\nFusce\r\naliquam,\r\nenim vel\r\nfermentu\r\nm ultrices,\r\nelit\r\nquam\r\nsemper erat,\r\nvitae\r\nsemper velit\r\naugue\r\nnon\r\nmagna.\r\n\r\nQuisque\r\nmaximu\r\ns semper\r\narcu, id\r\npellente\r\nsque est\r\ntempus\r\na.\r\nPhasellus lacus\r\nelit,\r\nauctor sit\r\namet\r\nlacinia a,\r\ndapibus\r\nvitae velit.\r\nPhasellus\r\nut\r\npharetra justo,\r\nut\r\nultricies erat.\r\nSed\r\nmolestie\r\nsapien vel\r\ninterdum\r\nlobortis.\r\nNulla\r\nfacilisi.\r\nVestibulum\r\nante\r\nipsum primis\r\nin\r\nfaucibus orci\r\nluctus et\r\nultrices\r\nposuere\r\ncubilia\r\ncurae;\r\nNulla nec\r\nmauris\r\nquis nisi\r\nvulputate\r\ngravida\r\nquis nec\r\nvelit.\r\n\r\nNam et\r\ncongue\r\nipsum.\r\nNulla vel\r\nelit non\r\ndolor\r\nmollis\r\naliquet vel at\r\nmagna.\r\nPellente\r\nsque nec\r\nfacilisis\r\nelit. In\r\nvulputate\r\nquis\r\nsem porta\r\nsuscipit.\r\nNullam\r\nsed ex\r\nornare\r\nnibh\r\nsuscipit mattis\r\nquis non\r\nlacus.\r\nMauris vel\r\nex urna.\r\nVivamus\r\nultricies\r\nsapien\r\nsit amet\r\nsapien\r\nvehicula\r\ngravida.\r\nDonec\r\nfeugiat\r\nvolutpat\r\nquam.\r\nVestibulum\r\nauctor\r\ndictum nisl,\r\nid\r\nhendrerit\r\nmetus\r\nullamcorper\r\nsed.\r\nNulla\r\nmaximus lacus\r\nvel\r\nmollis\r\nmaximus. Nulla\r\nlaoreet\r\nplacerat\r\nquam eu\r\nviverra.\r\nEtiam\r\nfeugiat\r\naccumsan\r\nnisl a\r\ncondiment\r\num. Sed\r\nultricies\r\nante ante,\r\nac\r\nauctor ligula\r\ngravida\r\nnec.\r\nPraesent a\r\nneque\r\ndignissim,\r\nsagittis\r\nfelis sit\r\namet,\r\ncondimentum\r\nturpis.\r\n\r\nFusce at\r\nleo vel\r\nest\r\nblandit\r\nmalesuada.\r\nPellentesqu\r\ne et\r\nneque non\r\nmetus\r\npellentesqu\r\ne\r\nimperdiet.\r\nPraesent\r\npellentesque\r\nlacinia\r\nlorem, et\r\ntristique\r\ntellus\r\nefficitur id.\r\nSuspend\r\nisse\r\naliquet\r\nultricies justo\r\nvitae\r\ninterdum.\r\nCras\r\ntristique\r\nviverra\r\nquam, eget\r\ngravida\r\nmi\r\nfermentum\r\nimperdiet.\r\nSed\r\nimperdiet\r\nvitae purus\r\nut\r\nvolutpat. Nulla\r\nlacinia\r\nelit in\r\nfermentum\r\nconsect\r\netur.\r\nPhasellus\r\ncommodo\r\nut nisl\r\nsit amet\r\nsagittis.\r\nDuis ac\r\nornare\r\norci.\r\n\r\nVivamus\r\nvel enim\r\nposuere,\r\npharetra\r\nex vel,\r\nelementu\r\nm est.\r\nVestibulum\r\ncommo\r\ndo luctus\r\nmetus\r\neget\r\nmaximus.\r\nSuspendis\r\nse a nulla\r\na odio\r\neleifend\r\nfaucibus.\r\nSuspend\r\nisse\r\nsemper\r\nlacus non\r\nporttitor\r\naliquet.\r\nCras ac\r\nscelerisqu\r\ne magna,\r\net\r\npulvinar justo.\r\nInteger\r\ncursus\r\npulvinar\r\nfringilla.\r\nMauris\r\nimperdiet\r\nnibh sit\r\namet\r\ntempor\r\nlaoreet. Morbi\r\ntincidun\r\nt tortor\r\nex, sit\r\namet\r\nmaximus purus\r\ntristique\r\nquis.\r\nQuisque\r\nsed\r\nhendrerit velit.\r\nMauris\r\nmattis\r\nnibh ut\r\neros luctus,\r\neget\r\nmattis\r\nmassa auctor.\r\nPhasellu\r\ns eu\r\nneque at\r\naugue\r\ngravida\r\nsagittis nec\r\nnon tortor.\r\nEtiam\r\nporttitor\r\nsem\r\nsodales mi\r\nullamcorp\r\ner\r\ngravida.\r\n\r\nIn in\r\ndictum orci.\r\nIn vitae\r\nvestibulu\r\nm quam.\r\nCras\r\naugue eros,\r\ntincidun\r\nt ac elit\r\nposuere,\r\nsollicitu\r\ndin\r\nefficitur\r\nlectus.\r\nPraesent quis\r\nsodales\r\nnisl. Proin\r\nsit amet\r\nmolestie\r\nest. In\r\ncommod\r\no mauris\r\nvel\r\nmauris\r\nefficitur, nec\r\nmollis\r\nmauris\r\nsagittis.\r\nCras ligula\r\nnibh,\r\negestas sit\r\namet eros\r\nin,\r\nlacinia\r\ntristique\r\nmagna. Cras\r\nrisus\r\nlibero, lacinia\r\neget\r\nlibero vitae,\r\nmaximu\r\ns aliquet\r\nnibh.\r\nMauris id\r\nsodales\r\npurus,\r\nvitae\r\ndictum\r\nlectus. Cras\r\nconsectetu\r\nr ligula\r\nvelit,\r\ntempus\r\npulvinar\r\nlacus\r\nporttitor vitae.\r\nPhasellu\r\ns eget\r\ntellus\r\nipsum.\r\n\r\nDonec\r\ninterdum\r\nlaoreet\r\nelit non\r\nvestibulum\r\n. Cras\r\nsed urna\r\nullamcorp\r\ner,\r\naliquam erat\r\neget,\r\nporta orci.\r\nVestibulu\r\nm eget\r\ncongue\r\nnulla. Sed\r\nsem\r\ntortor,\r\neuismod at\r\nrutrum id,\r\nsagittis a\r\nnunc.\r\nDuis in\r\nnibh\r\nfacilisis,\r\ndignissim\r\npurus ut,\r\nhendrerit\r\nmagna.\r\nSed\r\nsemper ligula\r\nid massa\r\nelement\r\num, non\r\nmalesua\r\nda velit\r\negestas.\r\nNullam\r\ndictum, mi\r\nnec\r\neuismod\r\nsagittis,\r\nligula leo\r\nullamcorp\r\ner dolor,\r\nquis\r\nfaucibus odio\r\nmetus\r\neget\r\nmagna. Ut\r\ngravida\r\nmetus non\r\nmetus\r\nbibendum\r\nbibendu\r\nm. In\r\nsagittis\r\neleifend\r\naliquet.\r\n\r\nInterdu\r\nm et\r\nmalesuada\r\nfames ac\r\nante\r\nipsum primis\r\nin\r\nfaucibus. Nam\r\nmollis\r\nsagittis\r\nfelis, in\r\nfaucibus\r\ntortor\r\npretium vel.\r\nNam\r\nnec enim\r\nmetus.\r\nDonec in\r\naugue\r\narcu. Proin\r\nnon\r\nlobortis\r\npurus, sit\r\namet\r\nlacinia elit.\r\nSuspendis\r\nse quis\r\neros\r\ncondimentum\r\n, blandit\r\njusto sit\r\namet,\r\nlobortis\r\nnisl.\r\nSuspendisse\r\nmaximu\r\ns massa\r\nsed urna\r\ntempor\r\nornare.\r\nNunc\r\nmalesuada\r\npurus\r\nodio, eu\r\nluctus\r\nlectus\r\nauctor nec.\r\nMorbi\r\nauctor\r\npellentesque\r\nauctor.\r\nSed\r\nullamcorper, ex\r\nvitae\r\naliquam\r\nvulputate,\r\nest diam\r\nfeugiat\r\nmi, id\r\nporttitor\r\nlectus orci\r\nac leo.\r\n\r\nDonec\r\nsit amet\r\nvelit\r\npulvinar,\r\nvenenatis\r\nturpis ut,\r\ninterdum\r\nligula.\r\nInterdum\r\net\r\nmalesuada\r\nfames ac\r\nante\r\nipsum primis\r\nin\r\nfaucibus.\r\nVestibulum\r\neu lacus\r\nurna.\r\nMaecenas\r\nsem\r\nnulla,\r\naccumsan eu\r\nultricies\r\nsed,\r\ntempor vel\r\nmagna.\r\nCras\r\naliquet\r\nsollicitudin\r\nsapien ac\r\npulvinar.\r\nPraesent\r\nac\r\nsodales mi.\r\nInteger\r\nvitae\r\nmauris massa.\r\nMaecen\r\nas iaculis\r\norci et\r\nfaucibus\r\ninterdum\r\n.\r\n\r\nNunc\r\nnec\r\nmaximus felis,\r\nsed\r\nfinibus\r\nquam.\r\nPellentesque\r\nfelis\r\nmassa,\r\nvestibulum in\r\ntellus\r\nvitae,\r\ncongue\r\ntincidunt justo.\r\nNunc\r\nvitae enim\r\nmalesua\r\nda,\r\nbibendum\r\nante nec,\r\nvarius\r\ntellus.\r\nPraesent vitae\r\nnisi id\r\nquam\r\nauctor\r\nlacinia at non\r\nquam.\r\nNam nec\r\nligula sit\r\namet\r\nfelis auctor\r\nsagittis.\r\nNunc in\r\nrisus eu\r\nurna\r\nvarius\r\nlaoreet quis\r\nsit amet\r\nfelis.\r\nMorbi varius\r\ntempor\r\norci, eu\r\nvestibulum\r\nnunc\r\nvestibulum\r\nac. Nunc\r\nvehicula\r\nvelit\r\neleifend\r\nconsequat\r\nporta.\r\nSuspendiss\r\ne\r\nmaximus\r\ndapibus orci, in\r\nvulputat\r\ne massa\r\npretium\r\nac.\r\nQuisque\r\nmalesuada\r\naliquet\r\naliquet.";
        
        /// <summary>
        /// Performance test for text wrapping. 
        /// Fixed kerning pairs major bottle-neck.
        /// </summary>
        [TestMethod, Ignore("This test should not run in a multithreaded test run. If we want to keep it, it should be moved to a separate benchmark project.")]
        [TestCategory("Benchmark")]
        public void Wrap20Paragraphs100Times()
        {
            List<string> longTexts = new List<string>();

            for (int i = 0; i < 100; i++)
            {
                longTexts.Add(LoremIpsum20Para);
            }

            var ttTextMeasurer = new FontMeasurerTrueType();
            ttTextMeasurer.SetFont(11d, "Aptos Narrow");
            double maxPixelWidth = 52d;

            Stopwatch timer = new Stopwatch();
            timer.Start();
            List<string> wrapped = new List<string>();
            foreach (string text in longTexts)
            {
                wrapped = ttTextMeasurer.MeasureAndWrapText(text, maxPixelWidth);
            }
            timer.Stop();

            Trace.WriteLine(timer.ElapsedMilliseconds);

            Assert.IsTrue(timer.ElapsedMilliseconds < 1200, "timer.ElapsedMilliseconds was > 1200, actual value: " + timer.ElapsedMilliseconds);

            ////Below is verification of previous text-wrapping.
            ////Might be unnecesary and can be removed in the future.
            ////Keep for now to ward against unintended text-wrap changes.
            ////////////////////////////////////////////////////////////////////
            //string outputStr = string.Join("\r\n", wrapped.ToArray());
            //File.WriteAllText("C:\\temp\\Optimized.txt", outputStr);

            //var currStr = File.ReadAllText("C:\\temp\\Optimized.txt");

            //List<string> differingStrings;
            //IEnumerable<string> ListNew = currStr.Split("\r\n").Distinct();
            //IEnumerable<string> ListPrev = UnOptimizedOriginalWrappingString.Split("\r\n").Distinct();

            //if (ListPrev.Count() > ListNew.Count())
            //    differingStrings = ListPrev.Except(ListNew).ToList();
            //else
            //    differingStrings = ListNew.Except(ListPrev).ToList();

            //Assert.AreEqual(0, differingStrings.Count());
            //Assert.AreEqual(UnOptimizedOriginalWrappingString, currStr);
        }

        [TestMethod]
        public void Wrap20Paragraphs100TimesMultipleTextFragments()
        {
            List<string> longTexts = new List<string>();
            List<MeasurementFont> fonts = new List<MeasurementFont>();

            MeasurementFont font = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11f,
                Style = MeasurementFontStyles.Regular
            };

            for (int i = 0; i < 100; i++)
            {
                longTexts.Add(LoremIpsum20Para);
                fonts.Add(font);
            }

            var ttTextMeasurer = new FontMeasurerTrueType();
            ttTextMeasurer.SetFont(11d, "Aptos Narrow");
            double maxPixelWidth = 52d;

            Stopwatch timer = new Stopwatch();
            timer.Start();
            List<string> wrapped = new List<string>();

            wrapped = ttTextMeasurer.WrapMultipleTextFragments(longTexts, fonts, maxPixelWidth);
            timer.Stop();

            Trace.WriteLine(timer.ElapsedMilliseconds);

            Assert.IsTrue(timer.ElapsedMilliseconds < 1000);

            ////Below is verification of previous text-wrapping.
            ////Might be unnecesary and can be removed in the future.
            ////Keep for now to ward against unintended text-wrap changes.
            ////////////////////////////////////////////////////////////////////
            //string outputStr = string.Join("\r\n", wrapped.ToArray());
            //File.WriteAllText("C:\\temp\\Optimized.txt", outputStr);

            //var currStr = File.ReadAllText("C:\\temp\\Optimized.txt");

            //List<string> differingStrings;
            //IEnumerable<string> ListNew = currStr.Split("\r\n").Distinct();
            //IEnumerable<string> ListPrev = UnOptimizedOriginalWrappingString.Split("\r\n").Distinct();

            //if (ListPrev.Count() > ListNew.Count())
            //    differingStrings = ListPrev.Except(ListNew).ToList();
            //else
            //    differingStrings = ListNew.Except(ListPrev).ToList();

            //Assert.AreEqual(0, differingStrings.Count());
            //Assert.AreEqual(UnOptimizedOriginalWrappingString, currStr);
        }
        [TestMethod]
        public void SettingBoldItalicAgainShouldNotTimeout()
        {
            MeasurementFont boldItalic = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11f,
                Style = MeasurementFontStyles.Bold | MeasurementFontStyles.Italic
            };

            var ttTextMeasurer = new FontMeasurerTrueType();

            Stopwatch timer = new Stopwatch();
            timer.Start();
            ttTextMeasurer.SetFont(boldItalic);
            timer.Stop();
            var firstTime = timer.ElapsedMilliseconds;

            timer.Restart();
            ttTextMeasurer.SetFont(boldItalic);
            timer.Stop();

            //Doing the same operation again should not be a whole second longer
            Assert.IsTrue((firstTime + 1000) > timer.ElapsedMilliseconds);
            //At time of writing OpenTypeFontCache.GetFromCache is 2s therefore it should take less
            Assert.IsTrue(timer.ElapsedMilliseconds < 2000);
        }

        [TestMethod]
        public void GetFromCacheBoldItalicShouldWork()
        {
            MeasurementFont boldItalic = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11f,
                Style = MeasurementFontStyles.Bold | MeasurementFontStyles.Italic
            };

            var ttTextMeasurer = new FontMeasurerTrueType();
            ttTextMeasurer.SetFont(boldItalic);

            var cachedFont = OpenTypeFontCache.GetFromCache("Aptos Narrow", FontSubFamily.BoldItalic);
            Assert.IsNotNull(cachedFont);
        }
    }
}
