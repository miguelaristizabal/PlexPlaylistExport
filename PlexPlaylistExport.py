"""This script exports Plex Playlists into the M3U format.

The script is designed in such a way that it only creates the M3U file
where the paths are altered in such a way that they are relative to the
'playlists' directory on the NAS and then link to the 'organized' folder
where all 'beets' music is located.
An example path would be '../organized/<artist>/<album>/<track>.ext'
Using beets we can then export the playlist for use on a USB thumbdrive by
using the 'beet move -e -d <dir> playlist:<playlistfile>.m3u' command.

Requirements
------------
  - plexapi: For communication with Plex
  - unidecode: To convert to ASCII codepage for backwards compatibility
"""

import argparse
import os
import re
import requests
import plexapi
import sys
import codecs
from plexapi.server import PlexServer
from unidecode import unidecode


def configure_stdio():
    """Use UTF-8 for console I/O so playlist names survive Windows code pages."""

    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="strict")

class ExportOptions():
    def __init__(self, args):
        self.host = args.host
        self.token = args.token
        self.export_all = args.all
        self.playlist = args.playlist
        self.asciify = args.asciify
        self.walkman = args.walkman
        self.writeAlbum = args.write_album
        self.writeAlbumArtist = args.write_album_artist
        self.plexMusicRoot = args.plex_music_root
        self.replaceWithDir = args.replace_with_dir
        self.outputDir = args.output_dir
        self.user = args.switch_user
        pass

def do_asciify(input):
    """ Converts a string to it's ASCII representation
    """
    
    if input == None:
        return None
    
    replaced = input
    replaced = replaced.replace('Ä', 'Ae')
    replaced = replaced.replace('ä', 'ae')
    replaced = replaced.replace('Ö', 'Oe')
    replaced = replaced.replace('ö', 'oe')
    replaced = replaced.replace('Ü', 'Ue')
    replaced = replaced.replace('ü', 'ue')
    replaced = unidecode(replaced)
    return replaced


def sanitize_filename(filename):
    """Make a title safe to use as a filename on Windows and Unix-like systems."""

    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
    sanitized = sanitized.rstrip(' .')
    if sanitized.upper() in {
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    }:
        sanitized = '_%s' % sanitized
    if sanitized == '':
        sanitized = 'playlist'
    return sanitized


def connect_to_plex(options: ExportOptions, status_stream):
    """Connect to Plex and optionally switch to a managed user."""

    print('Connecting to plex...', end='', file=status_stream, flush=True)
    try:
        plex = PlexServer(options.host, options.token)
    except (plexapi.exceptions.Unauthorized, requests.exceptions.ConnectionError):
        print(' failed', file=status_stream, flush=True)
        return None
    print(' done', file=status_stream, flush=True)

    if options.user != None:
        print('Switching to managed account %s...' % options.user, end='', file=status_stream, flush=True)
        try:
            plex = plex.switchUser(options.user)
        except (plexapi.exceptions.Unauthorized, requests.exceptions.ConnectionError):
            print(' failed', file=status_stream, flush=True)
            return None
        print(' done', file=status_stream, flush=True)

    return plex


def get_audio_playlists(plex: PlexServer, status_stream):
    """Fetch all audio playlists from Plex."""

    print('Getting playlists... ', end='', file=status_stream, flush=True)
    playlists = [item for item in plex.playlists() if item.playlistType == 'audio']
    print(' done', file=status_stream, flush=True)
    return playlists


def create_output_filename(playlist_title, extension, used_filenames=None):
    """Create a safe output filename and avoid collisions within a single run."""

    sanitized_title = sanitize_filename(playlist_title)
    filename = '%s.%s' % (sanitized_title, extension)
    if used_filenames == None:
        return filename

    candidate = filename
    suffix = 2
    while candidate in used_filenames:
        candidate = '%s (%s).%s' % (sanitized_title, suffix, extension)
        suffix += 1
    used_filenames.add(candidate)
    return candidate


def create_output_path(options: ExportOptions, playlist_title, extension, used_filenames=None):
    """Create the final output path for a playlist file."""

    filename = create_output_filename(playlist_title, extension, used_filenames)
    if options.outputDir != None:
        os.makedirs(options.outputDir, exist_ok=True)
        return os.path.join(options.outputDir, filename)
    return filename


def rewrite_media_path(filepath, options: ExportOptions):
    """Rewrite a Plex media path for playlist output."""

    if filepath.startswith(options.plexMusicRoot):
        suffix = filepath[len(options.plexMusicRoot):].lstrip('/\\')
        if options.walkman:
            return suffix

        replacement = options.replaceWithDir.rstrip('/\\')
        if replacement == '':
            return suffix
        if suffix == '':
            return replacement
        return '%s/%s' % (replacement, suffix)

    return filepath

def list_playlists(options: ExportOptions):
    """ Lists all 'audio' playlists on the given Plex server
    """

    plex = connect_to_plex(options, sys.stderr)
    if plex == None:
        return

    playlists = get_audio_playlists(plex, sys.stderr)

    print('', file=sys.stderr, flush=True)
    print('Supply any of the following playlists to --playlist <playlist>:', file=sys.stderr, flush=True)
    for item in playlists:
        # Print the playlist titles to stdout so they can be captured separately from the status outputs
        print('%s' % item.title, file=sys.stdout, flush=True)


def write_playlist_file(playlist, options: ExportOptions, used_filenames=None):
    """Write an M3U/M3U8 file for a playlist object."""

    playlist_title = do_asciify(playlist.title) if options.asciify else playlist.title
    extension = "m3u" if options.asciify else "m3u8"
    encoding = "ascii" if options.asciify else "utf-8"
    output_filename = create_output_path(options, playlist_title, extension, used_filenames)

    if os.path.basename(output_filename) != '%s.%s' % (playlist_title, extension) or options.outputDir != None:
        print('Writing playlist "%s" to "%s"...' % (playlist.title, output_filename))
    else:
        print('Writing playlist "%s"...' % playlist.title)

    m3u = open(output_filename, 'w', encoding=encoding)
    if not options.walkman:
        m3u.write('#EXTM3U\n')
        m3u.write('#PLAYLIST:%s\n' % playlist_title)
        m3u.write('\n')

    print('Iterating playlist...', end='')
    items = playlist.items()
    print(' %s items found' % playlist.leafCount)

    print('Writing M3U...', end='')
    for item in items:
        media = item.media[0]
        seconds = int(item.duration / 1000)
        title = do_asciify(item.title) if options.asciify else item.title
        album = do_asciify(item.parentTitle) if options.asciify else item.parentTitle
        artist = do_asciify(item.originalTitle) if options.asciify else item.originalTitle
        albumArtist = do_asciify(item.grandparentTitle) if options.asciify else item.grandparentTitle
        if artist == None:
            artist = albumArtist

        parts = media.parts
        if options.writeAlbum and not options.walkman:
            m3u.write('#EXTALB:%s\n' % album)
        if options.writeAlbumArtist and not options.walkman:
            m3u.write('#EXTART:%s\n' % albumArtist)
        for part in parts:
            rewritten_path = rewrite_media_path(part.file, options)
            if not options.walkman:
                m3u.write('#EXTINF:%s,%s - %s\n' % (seconds, artist, title))
            m3u.write('%s\n' % rewritten_path)
            if not options.walkman:
                m3u.write('\n')

    m3u.close()
    print(' done')

def export_playlist(options: ExportOptions):
    """ Exports a given playlist from the specified Plex server in M3U format.
    """

    plex = connect_to_plex(options, sys.stdout)
    if plex == None:
        return

    print('Getting playlist...', end='')
    try:
        playlist = plex.playlist(options.playlist)
    except (plexapi.exceptions.NotFound):
        print(' failed')
        return
    print(' done')

    write_playlist_file(playlist, options)


def export_all_playlists(options: ExportOptions):
    """Export all audio playlists in one run without shell-level title piping."""

    plex = connect_to_plex(options, sys.stdout)
    if plex == None:
        return

    playlists = get_audio_playlists(plex, sys.stdout)
    print('Exporting %s audio playlists...' % len(playlists))

    used_filenames = set()
    for playlist in playlists:
        try:
            write_playlist_file(playlist, options, used_filenames)
        except OSError as error:
            print(' failed to write "%s": %s' % (playlist.title, error))

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('-a', '--all',
        action = 'store_true',
        help = "Export all available audio playlists")
    mode.add_argument(
        '-p', '--playlist',
        type = str,
        help = "The name of the Playlist in Plex for which to create the M3U playlist"
    )
    mode.add_argument('-l', '--list',
        action = 'store_true',
        help = "Use this option to get a list of all available playlists")
    
    parser.add_argument(
        '--asciify',
        action = 'store_true',
        help = "If enabled, tries to ASCII-fy encountered Unicode characters. This can be important for backwards compatiblity with certain older hardware.\nIt only applies to #EXT<xxx> lines. Paths will need to be handled otherwise."
    )
    parser.add_argument(
        '--walkman',
        action = 'store_true',
        help = "Write a minimal playlist without comment lines and strip the plex music root from paths so the result is relative."
    )
    parser.add_argument(
        '--write-album',
        action = 'store_true',
        help = "If enabled, the playlist will include the Album title in separate #EXTALB lines"
    )
    parser.add_argument(
        '--write-album-artist',
        action = 'store_true',
        help = "If enabled, the playlist will include the Albumartist in separate #EXTART lines"
    )
    parser.add_argument(
        '--host',
        type = str,
        help = "The URL to the Plex Server, i.e.: http://192.168.0.100:32400",
        default = 'http://192.168.0.100:32400'
    )
    parser.add_argument(
        '--token',
        type = str,
        help = "The Token used to authenticate with the Plex Server",
        default = 'qwAUDPoVCf4x1KJ9GJbJ'
    )
    parser.add_argument(
        '--plex-music-root',
        type = str,
        help = "The root of the plex music library location, for instance '/music'",
        default = '/music'
    )
    parser.add_argument(
        '--replace-with-dir',
        type = str,
        help = "The string which we replace the plex music library root dir with in the M3U. This could be a relative path for instance '..'.",
        default = '..'
    )
    parser.add_argument(
        '--output-dir',
        type = str,
        help = "Optional: Directory where exported playlists will be written. Defaults to the current working directory."
    )
    parser.add_argument(
        '-u', '--switch-user',
        type = str,
        help = "Optional: The Managed User Account you want to switch to upon connect."
    )
    
    args = parser.parse_args()
    options = ExportOptions(args=args)

    if (args.list):
        list_playlists(options)
    elif (args.all):
        export_all_playlists(options)
    else:
        export_playlist(options)

if __name__ == "__main__":
    configure_stdio()
    main()
