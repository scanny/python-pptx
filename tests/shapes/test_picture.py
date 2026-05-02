"""Test suite for pptx.shapes.picture module."""

from __future__ import annotations

import io

import pytest

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.slide import SlidePart
from pptx.shapes.picture import Movie, Picture, _BasePicture, _MediaFormat
from pptx.util import Pt

from ..unitutil.cxml import element, xml
from ..unitutil.mock import call, class_mock, instance_mock, property_mock


def _sld_with_movie_cond(shape_id, cond_cxml):
    """Return a `p:sld` element containing a movie and a `p:cond` inside its timing node.

    *cond_cxml* is a cxml fragment like ``"p:cond{delay=indefinite}"`` that
    is rendered into the ``p:stCondLst`` of the video's ``p:cTn``.
    """
    cond_xml = xml(cond_cxml)
    # -- strip the outer xmlns decl that xml() adds so the fragment nests cleanly --
    cond_frag = cond_xml.split("?>", 1)[-1].strip()
    # -- rendered element has its own default-namespace decl; remove it so the
    #    resulting parse doesn't choke on duplicate ns prefixes --
    cond_frag = cond_frag.replace(
        ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"', ""
    )
    sld_xml = (
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        "  <p:cSld>"
        "    <p:spTree>"
        "      <p:pic>"
        "        <p:nvPicPr>"
        '          <p:cNvPr id="%d" name="m"/>'
        "          <p:cNvPicPr/>"
        "          <p:nvPr/>"
        "        </p:nvPicPr>"
        "      </p:pic>"
        "    </p:spTree>"
        "  </p:cSld>"
        "  <p:timing>"
        "    <p:tnLst>"
        "      <p:par>"
        '        <p:cTn id="1" nodeType="tmRoot">'
        "          <p:childTnLst>"
        "            <p:video>"
        "              <p:cMediaNode>"
        '                <p:cTn id="2">'
        "                  <p:stCondLst>"
        "                    %s"
        "                  </p:stCondLst>"
        "                </p:cTn>"
        "                <p:tgtEl>"
        '                  <p:spTgt spid="%d"/>'
        "                </p:tgtEl>"
        "              </p:cMediaNode>"
        "            </p:video>"
        "          </p:childTnLst>"
        "        </p:cTn>"
        "      </p:par>"
        "    </p:tnLst>"
        "  </p:timing>"
        "</p:sld>" % (shape_id, cond_frag, shape_id)
    )
    return parse_xml(sld_xml)


class Describe_BasePicture(object):
    def it_knows_its_cropping(self, crop_get_fixture):
        picture, prop_name, expected_value = crop_get_fixture
        crop = getattr(picture, prop_name)
        assert abs(crop - expected_value) < 0.000001

    def it_can_change_its_cropping(self, crop_set_fixture):
        picture, prop_name, value, expected_xml = crop_set_fixture
        setattr(picture, prop_name, value)
        assert picture._element.xml == expected_xml

    def it_provides_access_to_its_outline(self, line_fixture):
        picture = line_fixture
        line = picture.line
        assert isinstance(line, LineFormat)
        # exercise line to test has_line interface, .ln and .get_or_add_ln()
        picture.line.width = Pt(1)
        assert picture.line.width == Pt(1)

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:pic/p:blipFill", "left", 0.0),
            ("p:pic/p:blipFill/a:srcRect", "top", 0.0),
            ("p:pic/p:blipFill/a:srcRect{l=99999}", "bottom", 0.0),
            ("p:pic/p:blipFill/a:srcRect{l=42424}", "left", 0.42424),
            ("p:pic/p:blipFill/a:srcRect{t=-10000}", "top", -0.1),
            ("p:pic/p:blipFill/a:srcRect{r=250000}", "right", 2.5),
            ("p:pic/p:blipFill/a:srcRect{b=33333}", "bottom", 0.33333),
        ]
    )
    def crop_get_fixture(self, request):
        pic_cxml, side, expected_value = request.param
        picture = Picture(element(pic_cxml), None)
        prop_name = "crop_%s" % side
        return picture, prop_name, expected_value

    @pytest.fixture(
        params=[
            (
                "p:pic{a:b=c}/p:blipFill",
                "left",
                0.11,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{l=11000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill",
                "top",
                0.21,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{t=21000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill",
                "right",
                0.31,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{r=31000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill",
                "bottom",
                0.41,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{b=41000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{l=80000}",
                "left",
                0.21,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{l=21000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{t=70000}",
                "top",
                0.22,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{t=22000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{r=60000}",
                "right",
                0.23,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{r=23000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{b=50000}",
                "bottom",
                0.24,
                "p:pic{a:b=c}/p:blipFill/a:srcRect{b=24000}",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{l=90000}",
                "left",
                0,
                "p:pic{a:b=c}/p:blipFill/a:srcRect",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{t=91000}",
                "top",
                0.0,
                "p:pic{a:b=c}/p:blipFill/a:srcRect",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{r=92000}",
                "right",
                0,
                "p:pic{a:b=c}/p:blipFill/a:srcRect",
            ),
            (
                "p:pic{a:b=c}/p:blipFill/a:srcRect{b=93000}",
                "bottom",
                0.0,
                "p:pic{a:b=c}/p:blipFill/a:srcRect",
            ),
        ]
    )
    def crop_set_fixture(self, request):
        pic_cxml, side, value, expected_cxml = request.param
        pic = element(pic_cxml)
        picture = Picture(pic, None)
        prop_name = "crop_%s" % side
        expected_xml = xml(expected_cxml)
        return picture, prop_name, value, expected_xml

    @pytest.fixture
    def line_fixture(self):
        return _BasePicture(element("p:pic/p:spPr"), None)


class DescribeMovie(object):
    def it_knows_its_shape_type(self, shape_type_fixture):
        movie = shape_type_fixture
        assert movie.shape_type == MSO_SHAPE_TYPE.MEDIA

    def it_knows_its_media_type(self, media_type_fixture):
        movie = media_type_fixture
        assert movie.media_type == PP_MEDIA_TYPE.MOVIE

    def it_provides_access_to_its_media_format(self, format_fixture):
        movie, MediaFormat_, pic, parent, media_format_ = format_fixture
        media_format = movie.media_format
        MediaFormat_.assert_called_once_with(pic, parent)
        assert media_format is media_format_

    def it_provides_access_to_its_poster_frame_image(self, pfrm_fixture):
        movie, slide_part_, calls, expected_value = pfrm_fixture
        poster_frame = movie.poster_frame
        assert slide_part_.get_image.call_args_list == calls
        assert poster_frame == expected_value

    @pytest.mark.parametrize(
        ("cond_cxml", "expected_condition", "expected_start_time"),
        [
            # -- default PowerPoint encoding: onClick, no explicit delay --
            ('p:cond{delay=indefinite}', "onClick", None),
            # -- explicit evt="onClick" with numeric delay --
            ('p:cond{evt=onClick,delay=3000}', "onClick", 3.0),
            # -- no evt, delay=0 is PowerPoint "with previous" --
            ('p:cond{delay=0}', "withPrevious", 0.0),
            # -- no evt, numeric delay is with-previous + offset --
            ('p:cond{delay=1500}', "withPrevious", 1.5),
            # -- evt=onEnd is "after previous" --
            ('p:cond{evt=onEnd,delay=0}', "afterPrevious", 0.0),
            ('p:cond{evt=onEnd,delay=2500}', "afterPrevious", 2.5),
        ],
    )
    def it_reads_its_start_condition_and_start_time(
        self, cond_cxml, expected_condition, expected_start_time
    ):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml=cond_cxml)
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        assert movie.start_condition == expected_condition
        assert movie.start_time == expected_start_time

    def it_returns_defaults_when_no_timing_node_is_present(self):
        # -- a p:pic under p:sld/p:cSld/p:spTree but no p:timing --
        sld = parse_xml(
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            "  <p:cSld>"
            "    <p:spTree>"
            "      <p:pic>"
            "        <p:nvPicPr>"
            '          <p:cNvPr id="42" name="m"/>'
            "          <p:cNvPicPr/>"
            "          <p:nvPr/>"
            "        </p:nvPicPr>"
            "      </p:pic>"
            "    </p:spTree>"
            "  </p:cSld>"
            "</p:sld>"
        )
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        assert movie.start_condition == "onClick"
        assert movie.start_time is None

    @pytest.mark.parametrize(
        ("condition", "expected_evt", "expected_delay"),
        [
            ("onClick", None, "indefinite"),
            ("withPrevious", None, 0),
            ("afterPrevious", "onEnd", 0),
        ],
    )
    def it_can_change_its_start_condition(self, condition, expected_evt, expected_delay):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml="p:cond{delay=indefinite}")
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.start_condition = condition

        cond = sld.xpath(".//p:cond")[0]
        assert cond.evt == expected_evt
        assert cond.delay == expected_delay

    def it_preserves_a_numeric_start_time_when_start_condition_changes(self):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml="p:cond{delay=2500}")
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)
        assert movie.start_time == 2.5

        movie.start_condition = "afterPrevious"

        cond = sld.xpath(".//p:cond")[0]
        assert cond.evt == "onEnd"
        assert cond.delay == 2500
        assert movie.start_time == 2.5
        assert movie.start_condition == "afterPrevious"

    @pytest.mark.parametrize(
        ("value", "expected_delay"),
        [
            (0, 0),
            (1.5, 1500),
            (2, 2000),
            (None, "indefinite"),
        ],
    )
    def it_can_change_its_start_time(self, value, expected_delay):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml="p:cond{delay=indefinite}")
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.start_time = value

        cond = sld.xpath(".//p:cond")[0]
        assert cond.delay == expected_delay

    def it_raises_on_invalid_start_condition(self):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml="p:cond{delay=indefinite}")
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        with pytest.raises(ValueError, match="start_condition must be"):
            movie.start_condition = "nope"

    def it_raises_on_invalid_start_time(self):
        sld = _sld_with_movie_cond(shape_id=42, cond_cxml="p:cond{delay=indefinite}")
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        with pytest.raises(ValueError, match="start_time must be"):
            movie.start_time = -1.0

    def it_can_replace_its_media_blob(self, request):
        # -- construct a p:pic with a video media reference plus an extLst/p14:media --
        pic_xml = (
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="movie.mp4"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr>"
            '      <a:videoFile r:link="rId_old_v"/>'
            "      <p:extLst>"
            '        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
            '          <p14:media'
            '            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
            '            r:embed="rId_old_m"/>'
            "        </p:ext>"
            "      </p:extLst>"
            "    </p:nvPr>"
            "  </p:nvPicPr>"
            '  <p:blipFill><a:blip r:embed="rId_poster"/><a:stretch/></p:blipFill>'
            "  <p:spPr/>"
            "</p:pic>"
        )
        pic = parse_xml(pic_xml)
        movie = Movie(pic, None)
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_video_media_part.return_value = ("rId_new_m", "rId_new_v")
        property_mock(request, Movie, "part", return_value=slide_part_)

        movie.replace_media(io.BytesIO(b"new-video-bytes"), mime_type="video/mp4")

        # -- rewrote rIds on the pic XML --
        assert pic.media_video_rId == "rId_new_v"
        assert pic.media_embed_rId == "rId_new_m"
        # -- dropped old rels --
        assert slide_part_.drop_rel.call_args_list == [call("rId_old_v"), call("rId_old_m")]
        # -- called with a Video value object --
        (video_arg,), _ = slide_part_.get_or_add_video_media_part.call_args
        assert video_arg.content_type == "video/mp4"

    def it_preserves_the_audio_tag_on_replace_media(self, request):
        # -- a pic with an a:audioFile (audio clip added via add_movie) --
        pic_xml = (
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="clip.wav"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr>"
            '      <a:audioFile r:link="rId_a"/>'
            "      <p:extLst>"
            '        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
            '          <p14:media'
            '            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
            '            r:embed="rId_m"/>'
            "        </p:ext>"
            "      </p:extLst>"
            "    </p:nvPr>"
            "  </p:nvPicPr>"
            '  <p:blipFill><a:blip r:embed="rId_poster"/><a:stretch/></p:blipFill>'
            "  <p:spPr/>"
            "</p:pic>"
        )
        pic = parse_xml(pic_xml)
        movie = Movie(pic, None)
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.get_or_add_video_media_part.return_value = ("rId_m2", "rId_a2")
        property_mock(request, Movie, "part", return_value=slide_part_)

        movie.replace_media(io.BytesIO(b"payload"), mime_type="audio/mpeg")

        # -- the a:audioFile tag is preserved; we only swapped rIds --
        audioFiles = pic.xpath("./p:nvPicPr/p:nvPr/a:audioFile")
        assert len(audioFiles) == 1
        assert pic.media_video_rId == "rId_a2"
        assert pic.media_embed_rId == "rId_m2"

    def it_raises_when_replace_media_called_on_non_media_pic(self):
        # -- a plain p:pic (no a:videoFile / a:audioFile) should reject replace_media --
        pic = parse_xml(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="p"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr/>"
            "  </p:nvPicPr>"
            '  <p:blipFill><a:blip r:embed="rId_poster"/><a:stretch/></p:blipFill>'
            "  <p:spPr/>"
            "</p:pic>"
        )
        movie = Movie(pic, None)

        with pytest.raises(ValueError, match="not a media pic"):
            movie.replace_media(io.BytesIO(b"data"), mime_type="video/mp4")

    def it_adds_a_cond_element_when_stCondLst_is_empty_on_write(self):
        # -- a p:video exists but its stCondLst has no p:cond yet --
        sld = parse_xml(
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            "  <p:cSld>"
            "    <p:spTree>"
            "      <p:pic>"
            "        <p:nvPicPr>"
            '          <p:cNvPr id="42" name="m"/>'
            "          <p:cNvPicPr/>"
            "          <p:nvPr/>"
            "        </p:nvPicPr>"
            "      </p:pic>"
            "    </p:spTree>"
            "  </p:cSld>"
            "  <p:timing>"
            "    <p:tnLst>"
            "      <p:par>"
            '        <p:cTn id="1" nodeType="tmRoot">'
            "          <p:childTnLst>"
            "            <p:video>"
            "              <p:cMediaNode>"
            '                <p:cTn id="2"/>'
            "                <p:tgtEl>"
            '                  <p:spTgt spid="42"/>'
            "                </p:tgtEl>"
            "              </p:cMediaNode>"
            "            </p:video>"
            "          </p:childTnLst>"
            "        </p:cTn>"
            "      </p:par>"
            "    </p:tnLst>"
            "  </p:timing>"
            "</p:sld>"
        )
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.start_time = 1.0

        cond = sld.xpath(".//p:cond")[0]
        assert cond.delay == 1000

    def it_can_delete_itself_and_clean_up_its_rels_and_timing(
        self, part_prop_, slide_part_
    ):
        # -- issue #974: removing a movie used to corrupt the file because the
        # -- p:video timing entry and the three media-related rels were left
        # -- dangling. Verify the override drops all three rels and removes
        # -- the matching p:video. --
        part_prop_.return_value = slide_part_
        sld = parse_xml(
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:cSld>"
            "    <p:spTree>"
            "      <p:pic>"
            "        <p:nvPicPr>"
            '          <p:cNvPr id="42" name="m"/>'
            "          <p:cNvPicPr/>"
            "          <p:nvPr>"
            '            <a:videoFile r:link="rId3"/>'
            "            <p:extLst>"
            '              <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
            '                <p14:media xmlns:p14="http://schemas.microsoft.com/'
            'office/powerpoint/2010/main" r:embed="rId2"/>'
            "              </p:ext>"
            "            </p:extLst>"
            "          </p:nvPr>"
            "        </p:nvPicPr>"
            '        <p:blipFill><a:blip r:embed="rId4"/></p:blipFill>'
            "      </p:pic>"
            "    </p:spTree>"
            "  </p:cSld>"
            "  <p:timing>"
            "    <p:tnLst>"
            "      <p:par>"
            '        <p:cTn id="1" nodeType="tmRoot">'
            "          <p:childTnLst>"
            "            <p:video>"
            "              <p:cMediaNode>"
            '                <p:cTn id="2"/>'
            "                <p:tgtEl>"
            '                  <p:spTgt spid="42"/>'
            "                </p:tgtEl>"
            "              </p:cMediaNode>"
            "            </p:video>"
            "          </p:childTnLst>"
            "        </p:cTn>"
            "      </p:par>"
            "    </p:tnLst>"
            "  </p:timing>"
            "</p:sld>"
        )
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.delete()

        # -- p:pic gone from spTree --
        assert sld.xpath(".//p:pic") == []
        # -- p:video timing entry for this shape removed --
        assert sld.xpath(".//p:video") == []
        # -- all three media-related rels dropped --
        dropped = {c.args[0] for c in slide_part_.drop_rel.call_args_list}
        assert dropped == {"rId2", "rId3", "rId4"}

    def it_can_delete_an_audio_movie_and_drop_its_link_rel(
        self, part_prop_, slide_part_
    ):
        # -- audio clips use <a:audioFile> instead of <a:videoFile>; both should be
        # -- picked up by the rId-collection xpath. --
        part_prop_.return_value = slide_part_
        sld = parse_xml(
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:cSld>"
            "    <p:spTree>"
            "      <p:pic>"
            "        <p:nvPicPr>"
            '          <p:cNvPr id="7" name="a"/>'
            "          <p:cNvPicPr/>"
            "          <p:nvPr>"
            '            <a:audioFile r:link="rId11"/>'
            "          </p:nvPr>"
            "        </p:nvPicPr>"
            "        <p:blipFill/>"
            "      </p:pic>"
            "    </p:spTree>"
            "  </p:cSld>"
            "</p:sld>"
        )
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.delete()

        dropped = {c.args[0] for c in slide_part_.drop_rel.call_args_list}
        assert dropped == {"rId11"}

    def it_deletes_cleanly_when_no_timing_node_is_present(
        self, part_prop_, slide_part_
    ):
        # -- some pre-existing movies (loaded from files authored elsewhere) may
        # -- have no p:video timing entry; delete() must tolerate that. --
        part_prop_.return_value = slide_part_
        sld = parse_xml(
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:cSld>"
            "    <p:spTree>"
            "      <p:pic>"
            "        <p:nvPicPr>"
            '          <p:cNvPr id="5" name="m"/>'
            "          <p:cNvPicPr/>"
            "          <p:nvPr>"
            '            <a:videoFile r:link="rId9"/>'
            "          </p:nvPr>"
            "        </p:nvPicPr>"
            '        <p:blipFill><a:blip r:embed="rId8"/></p:blipFill>'
            "      </p:pic>"
            "    </p:spTree>"
            "  </p:cSld>"
            "</p:sld>"
        )
        pic = sld.xpath(".//p:pic")[0]
        movie = Movie(pic, None)

        movie.delete()

        assert sld.xpath(".//p:pic") == []
        dropped = {c.args[0] for c in slide_part_.drop_rel.call_args_list}
        assert dropped == {"rId8", "rId9"}

    def it_provides_access_to_its_media_blob_ext_and_content_type(
        self, request, part_prop_, slide_part_
    ):
        # -- issue #801: expose Movie.blob / .ext / .content_type mirroring
        # -- Picture.image.blob / .ext / .content_type for picture shapes. --
        pic = parse_xml(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="clip.mp4"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr>"
            '      <a:videoFile r:link="rId_v"/>'
            "      <p:extLst>"
            '        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
            '          <p14:media'
            '            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
            '            r:embed="rId_m"/>'
            "        </p:ext>"
            "      </p:extLst>"
            "    </p:nvPr>"
            "  </p:nvPicPr>"
            "  <p:spPr/>"
            "</p:pic>"
        )
        media_part_ = instance_mock(request, MediaPart)
        media_part_.blob = b"\x00\x00movie-bytes"
        media_part_.content_type = "video/mp4"
        media_part_.partname = PackURI("/ppt/media/media1.mp4")
        part_prop_.return_value = slide_part_
        slide_part_.related_part.return_value = media_part_

        movie = Movie(pic, None)

        assert movie.blob == b"\x00\x00movie-bytes"
        assert movie.content_type == "video/mp4"
        assert movie.ext == "mp4"
        # -- p14:media @r:embed is preferred over a:videoFile @r:link --
        slide_part_.related_part.assert_called_with("rId_m")

    def it_falls_back_to_the_videoFile_link_when_p14_media_is_absent(
        self, request, part_prop_, slide_part_
    ):
        # -- audio clips round-tripped from files written by other tools may have
        # -- just a:videoFile / a:audioFile with no p14:media descriptor;
        # -- blob / ext / content_type must still resolve via the link rel. --
        pic = parse_xml(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="clip.wav"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr>"
            '      <a:audioFile r:link="rId_a"/>'
            "    </p:nvPr>"
            "  </p:nvPicPr>"
            "  <p:spPr/>"
            "</p:pic>"
        )
        media_part_ = instance_mock(request, MediaPart)
        media_part_.blob = b"RIFFaudio"
        media_part_.content_type = "audio/wav"
        media_part_.partname = PackURI("/ppt/media/media2.wav")
        part_prop_.return_value = slide_part_
        slide_part_.related_part.return_value = media_part_

        movie = Movie(pic, None)

        assert movie.blob == b"RIFFaudio"
        assert movie.content_type == "audio/wav"
        assert movie.ext == "wav"
        slide_part_.related_part.assert_called_with("rId_a")

    def it_returns_None_when_no_media_rel_is_present(self):
        # -- a malformed pic with neither a:videoFile / a:audioFile nor p14:media
        # -- exposes None on every media accessor — never raises. --
        pic = parse_xml(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="p"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr/>"
            "  </p:nvPicPr>"
            "  <p:spPr/>"
            "</p:pic>"
        )
        movie = Movie(pic, None)

        assert movie.blob is None
        assert movie.content_type is None
        assert movie.ext is None

    def it_returns_None_when_related_part_is_not_a_MediaPart(
        self, request, part_prop_, slide_part_
    ):
        # -- defensive: if the rId resolves to some non-media part (corrupted file),
        # -- don't return junk bytes — return None.
        pic = parse_xml(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            '        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            '        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships">'
            "  <p:nvPicPr>"
            '    <p:cNvPr id="5" name="clip.mp4"/>'
            "    <p:cNvPicPr/>"
            "    <p:nvPr>"
            '      <a:videoFile r:link="rId_v"/>'
            "    </p:nvPr>"
            "  </p:nvPicPr>"
            "  <p:spPr/>"
            "</p:pic>"
        )
        not_a_media_part = instance_mock(request, ImagePart)
        part_prop_.return_value = slide_part_
        slide_part_.related_part.return_value = not_a_media_part

        movie = Movie(pic, None)

        assert movie.blob is None
        assert movie.content_type is None
        assert movie.ext is None

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def format_fixture(self, _MediaFormat_, media_format_):
        pic = element("p:pic")
        parent = movie = Movie(pic, None)
        return movie, _MediaFormat_, pic, parent, media_format_

    @pytest.fixture
    def media_type_fixture(self):
        return Movie(None, None)

    @pytest.fixture(
        params=[
            ("p:pic/p:blipFill/a:blip{r:embed=rId42}", True),
            ("p:pic/p:blipFill/a:blip", False),
        ]
    )
    def pfrm_fixture(self, request, slide_part_, image_, part_prop_):
        cxml, is_present = request.param
        movie = Movie(element(cxml), None)
        calls, expected_value = [], None
        part_prop_.return_value = slide_part_
        slide_part_.get_image.return_value = image_
        if is_present:
            calls.append(call("rId42"))
            expected_value = image_
        return movie, slide_part_, calls, expected_value

    @pytest.fixture
    def shape_type_fixture(self):
        return Movie(None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def _MediaFormat_(self, request, media_format_):
        return class_mock(request, "pptx.shapes.picture._MediaFormat", return_value=media_format_)

    @pytest.fixture
    def media_format_(self, request):
        return instance_mock(request, _MediaFormat)

    @pytest.fixture
    def part_prop_(self, request):
        return property_mock(request, Movie, "part")

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)


class DescribePicture(object):
    def it_knows_its_masking_shape(self, autoshape_get_fixture):
        picture, expected_value = autoshape_get_fixture
        auto_shape_type = picture.auto_shape_type
        assert auto_shape_type == expected_value

    def it_can_change_its_masking_shape(self, autoshape_set_fixture):
        picture, new_value, expected_xml = autoshape_set_fixture
        picture.auto_shape_type = new_value
        assert picture._element.xml == expected_xml

    def it_knows_its_shape_type(self, shape_type_fixture):
        picture = shape_type_fixture
        assert picture.shape_type == MSO_SHAPE_TYPE.PICTURE

    def it_provides_access_to_its_image(self, image_fixture):
        picture, slide_part_, rId, image_ = image_fixture
        image = picture.image
        slide_part_.get_image.assert_called_once_with(rId)
        assert image is image_

    def it_can_delete_itself_and_drop_its_image_rel(self, part_prop_, slide_part_):
        spTree = element("p:spTree/(p:sp,p:pic/p:blipFill/a:blip{r:embed=rId7},p:sp)")
        pic = spTree.xpath("p:pic")[0]
        picture = Picture(pic, None)

        picture.delete()

        slide_part_.drop_rel.assert_called_once_with("rId7")
        assert spTree.xpath("p:pic") == []
        assert len(spTree.xpath("p:sp")) == 2

    def it_can_delete_itself_when_it_has_no_image_rel(self, part_prop_, slide_part_):
        spTree = element("p:spTree/(p:sp,p:pic/p:blipFill/a:blip,p:sp)")
        pic = spTree.xpath("p:pic")[0]
        picture = Picture(pic, None)

        picture.delete()

        slide_part_.drop_rel.assert_not_called()
        assert spTree.xpath("p:pic") == []

    def it_can_replace_its_embedded_image(
        self, part_prop_, slide_part_, image_part_
    ):
        pic = element("p:pic/p:blipFill/a:blip{r:embed=rIdOld}")
        picture = Picture(pic, None)
        slide_part_.get_or_add_image_part.return_value = (image_part_, "rIdNew")

        picture.replace_image("new.png")

        slide_part_.get_or_add_image_part.assert_called_once_with("new.png")
        slide_part_.drop_rel.assert_called_once_with("rIdOld")
        blip = pic.xpath(".//a:blip")[0]
        assert blip.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        ) == "rIdNew"

    def it_leaves_the_rel_in_place_when_replacement_dedupes_to_same_image(
        self, part_prop_, slide_part_, image_part_
    ):
        # -- When the provided image matches the already-embedded one,
        # -- get_or_add_image_part returns the same rId. The relationship
        # -- should *not* be dropped in that case.
        pic = element("p:pic/p:blipFill/a:blip{r:embed=rIdSame}")
        picture = Picture(pic, None)
        slide_part_.get_or_add_image_part.return_value = (image_part_, "rIdSame")

        picture.replace_image("same.png")

        slide_part_.drop_rel.assert_not_called()
        blip = pic.xpath(".//a:blip")[0]
        assert blip.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        ) == "rIdSame"

    def it_raises_when_replacing_a_picture_with_no_embedded_image(
        self, part_prop_, slide_part_
    ):
        # -- malformed a:blip with no r:embed attribute --
        pic = element("p:pic/p:blipFill/a:blip")
        picture = Picture(pic, None)

        with pytest.raises(ValueError, match="no embedded image to replace"):
            picture.replace_image("new.png")

        slide_part_.get_or_add_image_part.assert_not_called()
        slide_part_.drop_rel.assert_not_called()

    def it_preserves_position_size_and_crop_across_replace_image(
        self, part_prop_, slide_part_, image_part_
    ):
        pic = element(
            "p:pic/("
            "p:blipFill/(a:blip{r:embed=rIdOld},a:srcRect{l=10000,t=20000,r=30000,b=40000}),"
            "p:spPr/(a:xfrm{rot=5400000}/(a:off{x=914400,y=457200},a:ext{cx=1828800,cy=1371600}),"
            "a:prstGeom{prst=rect}/a:avLst)"
            ")"
        )
        picture = Picture(pic, None)
        slide_part_.get_or_add_image_part.return_value = (image_part_, "rIdNew")

        picture.replace_image("new.png")

        # -- embedded image rId swapped --
        blip = pic.xpath(".//a:blip")[0]
        rembed_qn = (
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        )
        assert blip.get(rembed_qn) == "rIdNew"
        # -- crop preserved verbatim --
        srcRect = pic.xpath(".//a:srcRect")[0]
        assert srcRect.get("l") == "10000"
        assert srcRect.get("t") == "20000"
        assert srcRect.get("r") == "30000"
        assert srcRect.get("b") == "40000"
        # -- rotation, position and extent preserved --
        xfrm = pic.xpath(".//a:xfrm")[0]
        assert xfrm.get("rot") == "5400000"
        off = pic.xpath(".//a:off")[0]
        assert (off.get("x"), off.get("y")) == ("914400", "457200")
        ext = pic.xpath(".//a:ext")[0]
        assert (ext.get("cx"), ext.get("cy")) == ("1828800", "1371600")
        # -- masking preset geometry preserved --
        prstGeom = pic.xpath(".//a:prstGeom")[0]
        assert prstGeom.get("prst") == "rect"

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("", None),
            ("/a:prstGeom{prst=rect}", MSO_SHAPE.RECTANGLE),
            ("/a:prstGeom{prst=hexagon}", MSO_SHAPE.HEXAGON),
        ]
    )
    def autoshape_get_fixture(self, request):
        prstGeom_cxml, expected_value = request.param
        pic_cxml = "p:pic/p:spPr%s" % prstGeom_cxml
        picture = Picture(element(pic_cxml), None)
        return picture, expected_value

    @pytest.fixture(
        params=[
            (
                "p:pic/p:spPr/a:custGeom",
                MSO_SHAPE.RECTANGLE,
                "p:pic/p:spPr/a:prstGeom{prst=rect}",
            ),
            (
                "p:pic/p:spPr/a:prstGeom{prst=rect}",
                MSO_SHAPE.OVAL,
                "p:pic/p:spPr/a:prstGeom{prst=ellipse}",
            ),
            (
                "p:pic/p:spPr/a:prstGeom{prst=ellipse}",
                MSO_SHAPE.PIE,
                "p:pic/p:spPr/a:prstGeom{prst=pie}",
            ),
        ]
    )
    def autoshape_set_fixture(self, request):
        pic_cxml, new_value, expected_cxml = request.param
        picture = Picture(element(pic_cxml), None)
        expected_xml = xml(expected_cxml)
        return picture, new_value, expected_xml

    @pytest.fixture
    def image_fixture(self, part_prop_, slide_part_, image_):
        pic_cxml, rId = "p:pic/p:blipFill/a:blip{r:embed=rId42}", "rId42"
        picture = Picture(element(pic_cxml), None)
        slide_part_.get_image.return_value = image_
        return picture, slide_part_, rId, image_

    @pytest.fixture
    def shape_type_fixture(self):
        return Picture(None, None)

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def part_prop_(self, request, slide_part_):
        return property_mock(request, Picture, "part", return_value=slide_part_)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)
